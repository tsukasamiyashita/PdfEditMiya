# -*- coding: utf-8 -*-
import os, sys, io, base64, cv2, csv, time, json, random, gc, re
import urllib.request
import urllib.error
import urllib.parse
import threading
import concurrent.futures
import numpy as np
import fitz
from PIL import Image
import google.generativeai as genai

try:
    from openai import AzureOpenAI
except ImportError:
    AzureOpenAI = None

from common import (
    parse_row_data,
    merge_2d_arrays_horizontally,
    normalize_text,
    parse_azure_endpoint_info,
    diagnose_api_error,
    format_api_error_message,
    DEFAULT_AZURE_API_VERSION,
    AZURE_API_VERSION_CANDIDATES
)

from engines import (
    expand_crop_rect_for_intersecting_objects,
    save_table_data_to_file
)

# ==============================
# AIエンジン共通ヘルパー関数群
# ==============================
def _build_extraction_prompt(is_table_format, crop_regions, custom_prompt=""):
    if is_table_format:
        if crop_regions:
            prompt = """
あなたは優秀なデータ入力オペレーターです。添付された画像には、1つ以上のテキスト領域（抽出範囲）が上から下に向かって配置されています。
それぞれの領域のテキストを読み取り、以下のJSON形式で出力してください。

【特別ルール（絶対厳守）】
- 画像内に複数の領域がある場合、上から順番にそれぞれのデータを読み取ってください。
- 各領域内で意味的なデータの区切り（項目名と値、または別々の行など）がある場合は、改行で繋げずに、必ず配列（リスト）の別々の要素に分割（セル分け）して出力してください。
- 縦書きのテキストは、必ず繋げて横書きの1つの文字列に変換してください。
- 『〃』『...』『同上』などの同上を表す記号が記載されている場合は、空白扱いにせず、必ずその記号をそのまま出力してください。

【出力形式（絶対厳守）】
以下のように、画像内の領域の数と同じ要素数を持つ配列を `regions` キーの値としてJSONオブジェクトで出力してください。
各要素は、その領域から抽出された各データを格納した配列です。
{
  "regions": [
    [ "領域1のデータ1", "領域1のデータ2", "領域1のデータ3..." ],
    [ "領域2のデータ1..." ]
  ]
}
""".strip()
        else:
            prompt = """
あなたは優秀なデータ入力オペレーターです。添付された図面管理台帳などの表画像を読み取り、正確なJSONデータを作成してください。

【データの分離精度を最優先する特別ルール（超重要）】
- データの見た目や文脈の意味よりも、「表の縦の罫線」や「文字列間の大きな空白」などの【物理的な列の区切り】を絶対的な基準として最優先し、物理的に分かれているデータは必ず別の要素として分割してください。
- 意味的に繋がっているように見えても、罫線や空白で区切られていれば絶対に1つの要素に結合しないでください。
- データが存在しない「空白セル」の場合は無視せず、必ず `""` (空文字) を該当位置に挿入し、列の位置を厳密に揃えてください。ただし、『〃』『...』『同上』などの同上を表す記号が記載されている場合は、絶対に空白にせず、その記号をそのまま出力してください。
- 1行のデータを丸ごと1つの文字列に繋げて出力することは絶対に禁止します。必ず各セルごとにリスト内で分割してください。
- 縦書きのテキスト（文字が縦に並んでいるもの）は、1文字ずつ分割したり複数行に分けたりせず、必ず繋げて横書きの1つの文字列に変換して出力してください。

【出力形式（絶対厳守）】
シンプルな配列（リスト）を持つJSONオブジェクトで出力してください。
{
  "header": ["列1名", "列2名", "列3名"],
  "rows": [
    ["データ1", "データ2", "", "データ4"],
    ["データ1", "データ2", "データ3", ""]
  ]
}
""".strip()
    else:
        if crop_regions:
            prompt = f"""
添付された画像には上から順に {len(crop_regions)} 個の抽出領域が含まれています。各領域のテキストを読み取り、領域ごとに以下の形式でプレーンテキストとして出力してください。
縦書きの文章は改行を取り除き、横書きに変換して出力してください。
『〃』『...』『同上』などの同上を表す記号が記載されている場合は、必ずそのまま出力してください。
---領域1---
テキスト内容
---領域2---
テキスト内容
""".strip()
        else:
            prompt = "この画像に記載されている手書きの文字や文章を可能な限り正確に読み取り、プレーンテキストとして出力してください。また、縦書きの文章は改行を取り除き、横書きに変換して出力してください。『〃』『...』『同上』などの同上を表す記号が記載されている場合は、必ずそのまま出力してください。"

    if custom_prompt and custom_prompt.strip():
        prompt += f"\n\n【ユーザーからの追加指示（厳守すること）】\n{custom_prompt.strip()}"

    return prompt

def _prepare_page_combined_image(f_path, page_num, crop_regions, is_table_format):
    """PDFの指定ページから画像を作成し、crop_regionsがある場合は縦に結合したPIL Imageと各領域のshape情報を返す"""
    with fitz.open(f_path) as doc:
        pix = doc[page_num].get_pixmap(dpi=300)

    img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
    if pix.n == 4:
        img_array = cv2.cvtColor(img_array, cv2.COLOR_RGBA2RGB)
    elif pix.n == 1:
        img_array = cv2.cvtColor(img_array, cv2.COLOR_GRAY2RGB)

    combined_img = None
    cropped_info = []

    if crop_regions:
        h_img, w_img = img_array.shape[:2]
        cropped_images = []
        for region in crop_regions:
            rx1, ry1, rx2, ry2 = region[:4]
            is_vert = region[4] if len(region) > 4 else False
            is_line = is_vert or abs(ry2 - ry1) < 0.03 or abs(rx2 - rx1) < 0.03

            if not is_table_format or is_line:
                x1, y1, x2, y2 = expand_crop_rect_for_intersecting_objects(img_array, rx1, ry1, rx2, ry2)
            else:
                x1, y1 = int(min(rx1, rx2) * w_img), int(min(ry1, ry2) * h_img)
                x2, y2 = int(max(rx1, rx2) * w_img), int(max(ry1, ry2) * h_img)
                x1, y1 = max(0, x1), max(0, y1)
                x2, y2 = max(x1 + 1, min(w_img, x2)), max(y1 + 1, min(h_img, y2))

            crop = img_array[y1:y2, x1:x2]
            cropped_images.append(crop)
            cropped_info.append(crop.shape)

        if cropped_images:
            max_w = max(img.shape[1] for img in cropped_images)
            total_h = sum(img.shape[0] for img in cropped_images) + (len(cropped_images) - 1) * 20
            combined_array = np.full((total_h, max_w, 3), 255, dtype=np.uint8)

            current_y = 0
            for img in cropped_images:
                h_crop, w_crop = img.shape[:2]
                combined_array[current_y:current_y + h_crop, 0:w_crop] = img
                current_y += h_crop + 20

            combined_img = Image.fromarray(combined_array)
    else:
        combined_img = Image.fromarray(img_array)
        cropped_info = [img_array.shape]

    max_size = 2048
    if max(combined_img.width, combined_img.height) > max_size:
        ratio = max_size / max(combined_img.width, combined_img.height)
        new_size = (int(combined_img.width * ratio), int(combined_img.height * ratio))
        combined_img = combined_img.resize(new_size, Image.Resampling.LANCZOS)

    return combined_img, cropped_info

def _robust_json_parse(json_str):
    """データ量超過で途切れたJSONやMarkdownコードブロック付きJSONを修復してパースする"""
    s_raw = json_str.strip()
    # Markdownコードブロックの除去
    s_raw = re.sub(r'^```(?:json)?\s*', '', s_raw, flags=re.IGNORECASE)
    s_raw = re.sub(r'\s*```$', '', s_raw)
    json_match = re.search(r'\{.*\}', s_raw, re.DOTALL)
    if json_match:
        s_raw = json_match.group(0)

    try:
        return json.loads(s_raw)
    except json.JSONDecodeError as initial_err:
        try:
            s = s_raw.strip()
            quote_count = len(re.findall(r'(?<!\\)"', s))
            if quote_count % 2 != 0:
                s += '"'
            s = re.sub(r',\s*$', '', s)
            open_brackets = s.count('[') - s.count(']')
            open_braces = s.count('{') - s.count('}')
            if open_brackets > 0:
                s += ']' * open_brackets
            if open_braces > 0:
                s += '}' * open_braces
            return json.loads(s)
        except Exception:
            try:
                s2 = s_raw.strip()
                last_comma = s2.rfind(',')
                if last_comma != -1:
                    s2 = s2[:last_comma]
                    quote_count = len(re.findall(r'(?<!\\)"', s2))
                    if quote_count % 2 != 0:
                        s2 += '"'
                    ob = s2.count('[') - s2.count(']')
                    oc = s2.count('{') - s2.count('}')
                    if ob > 0:
                        s2 += ']' * ob
                    if oc > 0:
                        s2 += '}' * oc
                    return json.loads(s2)
            except Exception:
                pass
        raise initial_err

def _parse_and_save_page_result(extracted_text, success, last_error, is_table_format, crop_regions, cropped_info, page_num, total_p, base, digits, save_dir, out_format):
    """AI応答テキストをパースして最終テーブルを構築し、指定フォーマットでファイル保存する"""
    all_regions_data = []
    if success:
        if is_table_format:
            try:
                data = _robust_json_parse(extracted_text)

                if crop_regions:
                    parsed_regions = data.get("regions", []) if isinstance(data, dict) else (data if isinstance(data, list) else [])
                    if not isinstance(parsed_regions, list):
                        parsed_regions = [parsed_regions]

                    for reg_idx in range(len(crop_regions)):
                        if reg_idx < len(parsed_regions):
                            rows = parsed_regions[reg_idx]
                            if not isinstance(rows, list):
                                rows = [rows]
                        else:
                            rows = [f"領域{reg_idx+1}のデータ取得失敗"]

                        page_data_to_write = []
                        h_crop, w_crop = cropped_info[reg_idx][:2]
                        clean_rows = []
                        for r in rows:
                            val = " ".join([str(x) for x in r]) if isinstance(r, list) else str(r)
                            if '\n' in val:
                                lines = [l.strip() for l in val.split('\n') if l.strip()]
                                if all(len(l) <= 2 for l in lines):
                                    val = "".join(lines)
                            clean_rows.append(val)

                        if h_crop > w_crop * 1.5 and all(len(x.strip()) <= 2 for x in clean_rows if x.strip()):
                            page_data_to_write.append(["".join(clean_rows)])
                        else:
                            if clean_rows:
                                page_data_to_write.append(clean_rows)
                            else:
                                page_data_to_write.append([""])

                        all_regions_data.append(page_data_to_write)
                else:
                    header = data.get("header", []) if isinstance(data, dict) else []
                    rows = data.get("rows", []) if isinstance(data, dict) else []
                    if not header and not rows and isinstance(data, list):
                        if data:
                            header = data[0] if isinstance(data[0], list) else [str(data[0])]
                            rows = data[1:]
                    safe_header = [str(x).strip() for x in header] if isinstance(header, list) else []
                    page_col_count = len(safe_header)
                    for r in rows:
                        if isinstance(r, list) and len(r) > page_col_count:
                            page_col_count = len(r)
                    if not safe_header:
                        safe_header = [f"列{idx+1}" for idx in range(page_col_count)]
                    page_data_to_write = []
                    padded_header = (safe_header + [""] * page_col_count)[:page_col_count]
                    page_data_to_write.append(padded_header)
                    for row_data in rows:
                        parsed_r = parse_row_data(row_data)
                        safe_row_local = []
                        for val in (parsed_r + [""] * page_col_count)[:page_col_count]:
                            v_str = str(val)
                            if '\n' in v_str:
                                lines = [l.strip() for l in v_str.split('\n') if l.strip()]
                                if len(lines) > 1 and all(len(l) <= 2 for l in lines):
                                    v_str = "".join(lines)
                            safe_row_local.append(v_str)
                        if any(v != "" for v in safe_row_local):
                            page_data_to_write.append(safe_row_local)
                    all_regions_data.append(page_data_to_write)
            except Exception as e:
                for _ in range(len(crop_regions) if crop_regions else 1):
                    all_regions_data.append([[f"JSONパースエラー (データ量超過の可能性): {e}"]])
        else:
            if crop_regions:
                # ---領域X--- の見出しで正確に領域番号ごとに取得する（空領域があってもズレないようにする）
                matches = list(re.finditer(r'---領域(\d+)---\s*(.*?)(?=---領域\d+---|\Z)', extracted_text.strip(), re.DOTALL))
                if matches:
                    region_map = {}
                    for m in matches:
                        idx_num = int(m.group(1)) - 1
                        region_map[idx_num] = m.group(2).strip()
                    for reg_idx in range(len(crop_regions)):
                        val = region_map.get(reg_idx, "")
                        if val:
                            lines = [l.strip() for l in val.split('\n') if l.strip()]
                            all_regions_data.append([["\n".join(lines)]] if lines else [[""]])
                        else:
                            all_regions_data.append([[""]])
                else:
                    # タグなしで返ってきた場合のフォールバック
                    lines = [l.strip() for l in extracted_text.strip().split('\n') if l.strip()]
                    for reg_idx in range(len(crop_regions)):
                        if reg_idx < len(lines):
                            all_regions_data.append([[lines[reg_idx]]])
                        else:
                            all_regions_data.append([[""]])
            else:
                lines = [line.strip() for line in extracted_text.strip().split('\n') if line.strip()]
                all_regions_data.append([lines] if lines else [[""]])
    else:
        for _ in range(len(crop_regions) if crop_regions else 1):
            all_regions_data.append([[f"AI抽出失敗: {last_error}"]])

    # ページデータの統合とファイル保存
    if not is_table_format:
        if crop_regions:
            merged_row = []
            for region_data in all_regions_data:
                region_texts = []
                for r in region_data:
                    for c in r:
                        if str(c).strip():
                            region_texts.append(str(c).strip())
                merged_row.append("\n".join(region_texts) if region_texts else "")
            merged_data = [merged_row] if merged_row else [[""]]
        else:
            merged_row = []
            for region_data in all_regions_data:
                for r in region_data:
                    for c in r:
                        if str(c).strip():
                            merged_row.append(str(c).strip())
            merged_data = [merged_row] if merged_row else [[""]]
    else:
        merged_data = merge_2d_arrays_horizontally(all_regions_data)

    final_data = []
    page_info_str = f"{page_num+1}/{total_p}"

    if not is_table_format:
        max_cols = max((len(r) for r in merged_data), default=1)
        if crop_regions:
            header = ["ページ番号"] + [f"抽出範囲{i+1}" for i in range(max_cols)]
        else:
            header = ["ページ番号"] + [f"テキスト{i+1}" for i in range(max_cols)]
        final_data.append(header)
        for row in merged_data:
            final_data.append([page_info_str] + row)
    elif crop_regions:
        header = ["ページ番号"]
        if all_regions_data and merged_data:
            for i, region_data in enumerate(all_regions_data):
                num_cols = len(region_data[0]) if region_data and len(region_data) > 0 else 1
                if num_cols == 1:
                    header.append(f"抽出範囲{i+1}")
                else:
                    for j in range(num_cols):
                        header.append(f"抽出範囲{i+1}-{j+1}")
        else:
            header = ["ページ番号"] + [f"抽出範囲{idx+1}" for idx in range(len(crop_regions))]

        final_data.append(header)
        for row in merged_data:
            final_data.append([page_info_str] + row)
    else:
        for r_idx, row in enumerate(merged_data):
            if r_idx == 0:
                final_data.append(["ページ番号"] + row)
            else:
                final_data.append([page_info_str] + row)

    save_path = os.path.join(save_dir, f"{base}_Page_{str(page_num+1).zfill(digits)}_AI抽出")
    sheet_title = f"Page_{str(page_num+1).zfill(digits)}"
    save_table_data_to_file(final_data, save_path, out_format, sheet_title=sheet_title)


# ==============================
# Gemini API データ抽出タスク
# ==============================
def extract_gemini_task(files, save_dir, options, ui):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files:
        raise Exception("PDFファイルが含まれていません。")
    api_key = options.get("api_key", "")
    genai.configure(api_key=api_key)

    models_to_try = options.get("models_to_try", ["gemini-2.5-flash"])
    out_format = options.get("out_format", "xlsx")
    crop_regions = options.get("crop_regions", [])
    extract_mode = options.get("extract_mode", "table")

    is_table_format = out_format in ["csv", "xlsx", "json", "md", "docx"] and extract_mode == "table"

    api_plan = options.get("api_plan", "free")
    api_rpm = options.get("api_rpm", 12)
    temperature = options.get("temperature", 0.0)
    disable_safety = options.get("disable_safety", True)
    max_tokens = options.get("max_tokens", 8192)
    custom_prompt = options.get("custom_prompt", "")
    threads_setting = options.get("threads", 1 if api_plan == "free" else 5)

    prompt = _build_extraction_prompt(is_table_format, crop_regions, custom_prompt)

    generation_config = {
        "temperature": temperature,
        "max_output_tokens": max_tokens,
    }
    if is_table_format:
        generation_config["response_mime_type"] = "application/json"

    safety_settings = None
    if disable_safety:
        safety_settings = [
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"}
        ]

    MAX_RPM = api_rpm if api_rpm > 0 else 1
    max_workers = max(1, threads_setting)

    request_timestamps = []
    rate_limit_lock = threading.Lock()
    progress_lock = threading.Lock()
    completed_pages = 0
    shared_context = {"fatal_error": None}

    page_tasks = []
    ui.set_indeterminate("処理対象のページをスキャン中...")
    for f in files:
        if ui.is_cancelled():
            return
        try:
            with fitz.open(f) as doc:
                total_pages = len(doc)
            for page_num in range(total_pages):
                page_tasks.append({
                    "file_path": f,
                    "page_num": page_num,
                    "total_pages": total_pages
                })
        except Exception as e:
            print(f"Failed to scan {f}: {e}")
            continue

    total_tasks = len(page_tasks)
    if total_tasks == 0:
        return

    ui.update_overall(0, total_tasks, f"全体の進捗 ( 0 / {total_tasks} ページ完了 )")

    def process_single_page_task(task_info):
        if shared_context["fatal_error"] or ui.is_cancelled():
            return

        f_path = task_info["file_path"]
        page_num = task_info["page_num"]
        total_p = task_info["total_pages"]
        base = os.path.splitext(os.path.basename(f_path))[0]
        digits = max(2, len(str(total_p)))

        try:
            combined_img, cropped_info = _prepare_page_combined_image(f_path, page_num, crop_regions, is_table_format)

            max_retries = 8
            extracted_text, success, last_error = "", False, ""

            for attempt in range(max_retries):
                if shared_context["fatal_error"]:
                    raise Exception(shared_context["fatal_error"])
                if ui.is_cancelled():
                    return
                required_sleep = 0.0
                is_fatal_quota = False

                for model_name in models_to_try:
                    if shared_context["fatal_error"]:
                        raise Exception(shared_context["fatal_error"])
                    if ui.is_cancelled():
                        return
                    total_wait = 0.0

                    with rate_limit_lock:
                        now = time.time()
                        nonlocal request_timestamps
                        request_timestamps = [t for t in request_timestamps if now - t < 60.0]
                        min_interval = 1.5

                        if len(request_timestamps) >= MAX_RPM:
                            expected_run_time = request_timestamps[0] + 60.0
                            total_wait = expected_run_time - now
                        else:
                            if request_timestamps:
                                last_planned_time = max(request_timestamps)
                                if now < last_planned_time + min_interval:
                                    expected_run_time = last_planned_time + min_interval
                                    total_wait = expected_run_time - now
                                else:
                                    expected_run_time = now
                                    total_wait = 0.0
                            else:
                                expected_run_time = now
                                total_wait = 0.0
                        request_timestamps.append(expected_run_time)
                        request_timestamps.sort()

                    if total_wait > 0:
                        if total_wait > 2.0:
                            ui.set_indeterminate(f"通信間隔調整中... (約 {int(total_wait)}秒待機)")
                        wait_step = 0.5
                        while total_wait > 0:
                            if shared_context["fatal_error"]:
                                raise Exception(shared_context["fatal_error"])
                            if ui.is_cancelled():
                                return
                            time.sleep(min(wait_step, total_wait))
                            total_wait -= wait_step

                    ui.set_indeterminate(f"Gemini AI解析中... ( P.{page_num+1} )")

                    try:
                        model = genai.GenerativeModel(model_name)
                        contents = [prompt, combined_img]
                        if generation_config:
                            response = model.generate_content(contents, generation_config=generation_config, safety_settings=safety_settings)
                        else:
                            response = model.generate_content(contents, safety_settings=safety_settings)

                        if not response.parts:
                            raise Exception("安全フィルタ等によりブロックされました。")
                        extracted_text = normalize_text(response.text.strip())
                        success = True
                        break
                    except Exception as api_err:
                        err_str = str(api_err)
                        last_error = err_str
                        diag = diagnose_api_error("gemini", api_err, model_or_deployment=model_name)
                        if diag["kind"] in ["auth", "forbidden", "not_found"]:
                            msg = format_api_error_message(diag)
                            shared_context["fatal_error"] = msg
                            raise Exception(msg)
                        elif diag["kind"] == "rate_limit":
                            if "limit: 0" in err_str.lower() or "limit:0" in err_str.lower() or "per day" in err_str.lower() or "perday" in err_str.lower():
                                msg = format_api_error_message(diag)
                                shared_context["fatal_error"] = msg
                                raise Exception(msg)
                            m = re.search(r'retry in ([\d\.]+)s', err_str, re.IGNORECASE | re.DOTALL)
                            if not m:
                                m = re.search(r'seconds:\s*(\d+)', err_str, re.IGNORECASE | re.DOTALL)
                            if m:
                                wait_sec = float(m.group(1))
                                if wait_sec > 15.0:
                                    msg = format_api_error_message(diag)
                                    shared_context["fatal_error"] = msg
                                    raise Exception(msg)
                                required_sleep = wait_sec + 2.0
                            else:
                                if attempt >= 2:
                                    msg = format_api_error_message(diag)
                                    shared_context["fatal_error"] = msg
                                    raise Exception(msg)
                                required_sleep = 4.0 + random.uniform(1.0, 3.0)
                        continue

                if success or is_fatal_quota:
                    break
                if shared_context["fatal_error"]:
                    raise Exception(shared_context["fatal_error"])
                if ui.is_cancelled():
                    return

                if required_sleep > 0:
                    sleep_time = required_sleep
                else:
                    sleep_time = min(60.0, 4.0 * (2 ** attempt)) + random.uniform(1.0, 3.0)

                wait_step = 0.5
                while sleep_time > 0:
                    if shared_context["fatal_error"]:
                        raise Exception(shared_context["fatal_error"])
                    if ui.is_cancelled():
                        return
                    time.sleep(min(wait_step, sleep_time))
                    sleep_time -= wait_step

            _parse_and_save_page_result(
                extracted_text, success, last_error,
                is_table_format, crop_regions, cropped_info,
                page_num, total_p, base, digits, save_dir, out_format
            )
            gc.collect()

        except Exception as e:
            if shared_context["fatal_error"]:
                raise Exception(shared_context["fatal_error"])
            raise Exception(f"ページ {page_num+1} の処理中にエラーが発生しました: {e}")

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(process_single_page_task, task) for task in page_tasks]

        for future in concurrent.futures.as_completed(futures):
            if ui.is_cancelled() or shared_context["fatal_error"]:
                executor.shutdown(wait=False, cancel_futures=True)
                if shared_context["fatal_error"]:
                    raise Exception(shared_context["fatal_error"])
                return

            try:
                future.result()
            except Exception as e:
                executor.shutdown(wait=False, cancel_futures=True)
                raise e

            with progress_lock:
                completed_pages += 1
                ui.update_overall(completed_pages, total_tasks, f"全体の進捗 ( {completed_pages} / {total_tasks} ページ完了 )")

    ui.set_determinate(total_tasks, total_tasks, "すべての処理が完了しました")


# ==============================
# Azure OpenAI Service API 通信・データ抽出タスク
# ==============================
def _encode_pil_image_to_base64(pil_img):
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode("utf-8")

def _call_azure_openai_api(api_key, endpoint, deployment, api_version, prompt, pil_img=None, is_table_format=False, temperature=0.0, max_tokens=4096):
    """
    Azure OpenAI Service に対してリクエストを送信し、応答テキストを返す。
    openai パッケージが導入されている場合は公式 SDK (AzureOpenAI) を使用し、
    未導入環境では標準ライブラリ (urllib.request) による REST API 通信へ自動フォールバックする。
    また、推論モデル(o1/o3/o4等)と標準モデル(gpt-4o/gpt-4.1等)のパラメータ差異を自動吸収する。
    """
    base_endpoint, resolved_dep, resolved_ver = parse_azure_endpoint_info(endpoint, deployment, api_version)
    if not base_endpoint:
        raise Exception("エンドポイントURLが正しく入力されていません。")
    if not resolved_dep:
        raise Exception("デプロイ名が入力されていません。")
    if not api_key:
        raise Exception("APIキーが入力されていません。")

    if pil_img is not None:
        b64_img = _encode_pil_image_to_base64(pil_img)
        user_content = [
            {"type": "text", "text": prompt},
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:image/png;base64,{b64_img}",
                    "detail": "high"
                }
            }
        ]
    else:
        user_content = prompt

    messages = [
        {"role": "system", "content": "あなたは高精度な文書解析・データ入力アシスタントです。指示された出力形式を厳守して回答してください。"},
        {"role": "user", "content": user_content}
    ]

    dep_lower = resolved_dep.lower()
    is_reasoning_model = any(dep_lower.startswith(prefix) or f"-{prefix}" in dep_lower for prefix in ["o1", "o3", "o4", "gpt-5", "gpt-6"])

    use_max_completion_tokens = is_reasoning_model
    use_temperature = not is_reasoning_model
    use_json_format = is_table_format

    for param_attempt in range(4):
        kwargs = {
            "model": resolved_dep,
            "messages": messages,
        }
        if use_max_completion_tokens:
            kwargs["max_completion_tokens"] = max_tokens
        else:
            kwargs["max_tokens"] = max_tokens

        if use_temperature:
            kwargs["temperature"] = temperature

        if use_json_format:
            kwargs["response_format"] = {"type": "json_object"}

        try:
            if AzureOpenAI is not None:
                client = AzureOpenAI(
                    api_key=api_key,
                    api_version=resolved_ver,
                    azure_endpoint=base_endpoint,
                    max_retries=0,
                    timeout=120.0
                )
                response = client.chat.completions.create(**kwargs)
                if not response.choices or response.choices[0].message.content is None:
                    raise Exception("コンテンツフィルタ等により空の応答が返されました。")
                return response.choices[0].message.content.strip()
            else:
                # 標準ライブラリ urllib.request による REST API フォールバック
                url = (
                    f"{base_endpoint.rstrip('/')}/openai/deployments/"
                    f"{urllib.parse.quote(resolved_dep)}/chat/completions"
                    f"?api-version={urllib.parse.quote(resolved_ver)}"
                )
                payload = {k: v for k, v in kwargs.items() if k != "model"}
                req_data = json.dumps(payload).encode("utf-8")
                req = urllib.request.Request(
                    url,
                    data=req_data,
                    headers={
                        "Content-Type": "application/json",
                        "api-key": api_key
                    },
                    method="POST"
                )
                try:
                    with urllib.request.urlopen(req, timeout=120.0) as resp:
                        resp_body = json.loads(resp.read().decode("utf-8"))
                except urllib.error.HTTPError as http_err:
                    err_body = ""
                    try:
                        err_body = http_err.read().decode("utf-8", errors="replace")
                    except Exception:
                        pass
                    raise Exception(f"HTTP {http_err.code}: {err_body or http_err.reason}")

                choices = resp_body.get("choices", [])
                if not choices or choices[0].get("message", {}).get("content") is None:
                    raise Exception("コンテンツフィルタ等により空の応答が返されました。")
                return choices[0]["message"]["content"].strip()

        except Exception as e:
            err_str = str(e)
            err_lower = err_str.lower()
            # モデルごとのパラメータ非対応エラー(400)を自動検知して調整・再試行
            if "400" in err_str or "unsupported" in err_lower or "invalid_request_error" in err_lower:
                changed = False
                if "max_tokens" in err_lower and not use_max_completion_tokens:
                    use_max_completion_tokens = True
                    changed = True
                elif "max_completion_tokens" in err_lower and use_max_completion_tokens:
                    use_max_completion_tokens = False
                    changed = True
                if "temperature" in err_lower and use_temperature:
                    use_temperature = False
                    changed = True
                if "response_format" in err_lower and use_json_format:
                    use_json_format = False
                    changed = True
                if changed:
                    continue
            raise e

def test_azure_connection(api_key, endpoint, deployment, api_version):
    """設定ダイアログの「テスト」ボタンから呼ばれる接続確認関数"""
    reply = _call_azure_openai_api(
        api_key=api_key,
        endpoint=endpoint,
        deployment=deployment,
        api_version=api_version,
        prompt=f"Connection Test ({time.time()}): OK とだけ返答してください。",
        pil_img=None,
        is_table_format=False,
        temperature=0.0,
        max_tokens=32
    )
    return reply

def extract_azure_task(files, save_dir, options, ui):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files:
        raise Exception("PDFファイルが含まれていません。")

    azure_api_key = options.get("azure_api_key", "")
    azure_endpoint = options.get("azure_endpoint", "")
    azure_deployment = options.get("azure_deployment", "")
    azure_api_version = options.get("azure_api_version", "2024-10-21")

    base_endpoint, resolved_dep, resolved_ver = parse_azure_endpoint_info(
        azure_endpoint, azure_deployment, azure_api_version
    )
    if not azure_api_key or not base_endpoint or not resolved_dep:
        raise Exception("Azure OpenAI の接続情報（エンドポイント、APIキー、デプロイ名）が不足しています。\n「⚙️ 詳細設定」から設定を確認してください。")

    out_format = options.get("out_format", "xlsx")
    crop_regions = options.get("crop_regions", [])
    extract_mode = options.get("extract_mode", "table")

    is_table_format = out_format in ["csv", "xlsx", "json", "md", "docx"] and extract_mode == "table"

    api_rpm = options.get("azure_rpm", 60)
    temperature = options.get("azure_temperature", 0.0)
    max_tokens = options.get("azure_max_tokens", 4096)
    custom_prompt = options.get("azure_custom_prompt", "")
    threads_setting = options.get("azure_threads", 3)

    prompt = _build_extraction_prompt(is_table_format, crop_regions, custom_prompt)

    MAX_RPM = api_rpm if api_rpm > 0 else 1
    max_workers = max(1, threads_setting)

    request_timestamps = []
    rate_limit_lock = threading.Lock()
    progress_lock = threading.Lock()
    completed_pages = 0
    shared_context = {"fatal_error": None}

    page_tasks = []
    ui.set_indeterminate("処理対象のページをスキャン中...")
    for f in files:
        if ui.is_cancelled():
            return
        try:
            with fitz.open(f) as doc:
                total_pages = len(doc)
            for page_num in range(total_pages):
                page_tasks.append({
                    "file_path": f,
                    "page_num": page_num,
                    "total_pages": total_pages
                })
        except Exception as e:
            print(f"Failed to scan {f}: {e}")
            continue

    total_tasks = len(page_tasks)
    if total_tasks == 0:
        return

    ui.update_overall(0, total_tasks, f"全体の進捗 ( 0 / {total_tasks} ページ完了 )")

    def process_single_page_task(task_info):
        if shared_context["fatal_error"] or ui.is_cancelled():
            return

        f_path = task_info["file_path"]
        page_num = task_info["page_num"]
        total_p = task_info["total_pages"]
        base = os.path.splitext(os.path.basename(f_path))[0]
        digits = max(2, len(str(total_p)))

        try:
            combined_img, cropped_info = _prepare_page_combined_image(f_path, page_num, crop_regions, is_table_format)

            max_retries = 6
            extracted_text, success, last_error = "", False, ""

            for attempt in range(max_retries):
                if shared_context["fatal_error"]:
                    raise Exception(shared_context["fatal_error"])
                if ui.is_cancelled():
                    return

                required_sleep = 0.0
                total_wait = 0.0

                with rate_limit_lock:
                    now = time.time()
                    nonlocal request_timestamps
                    request_timestamps = [t for t in request_timestamps if now - t < 60.0]
                    min_interval = 0.5

                    if len(request_timestamps) >= MAX_RPM:
                        expected_run_time = request_timestamps[0] + 60.0
                        total_wait = expected_run_time - now
                    else:
                        if request_timestamps:
                            last_planned_time = max(request_timestamps)
                            if now < last_planned_time + min_interval:
                                expected_run_time = last_planned_time + min_interval
                                total_wait = expected_run_time - now
                            else:
                                expected_run_time = now
                                total_wait = 0.0
                        else:
                            expected_run_time = now
                            total_wait = 0.0
                    request_timestamps.append(expected_run_time)
                    request_timestamps.sort()

                if total_wait > 0:
                    if total_wait > 2.0:
                        ui.set_indeterminate(f"通信間隔調整中... (約 {int(total_wait)}秒待機)")
                    wait_step = 0.5
                    while total_wait > 0:
                        if shared_context["fatal_error"]:
                            raise Exception(shared_context["fatal_error"])
                        if ui.is_cancelled():
                            return
                        time.sleep(min(wait_step, total_wait))
                        total_wait -= wait_step

                ui.set_indeterminate(f"Azure OpenAI 解析中... ( P.{page_num+1} )")

                try:
                    raw_res = _call_azure_openai_api(
                        api_key=azure_api_key,
                        endpoint=base_endpoint,
                        deployment=resolved_dep,
                        api_version=resolved_ver,
                        prompt=prompt,
                        pil_img=combined_img,
                        is_table_format=is_table_format,
                        temperature=temperature,
                        max_tokens=max_tokens
                    )
                    extracted_text = normalize_text(raw_res)
                    success = True
                    break
                except Exception as api_err:
                    err_str = str(api_err)
                    last_error = err_str
                    err_lower = err_str.lower()
                    diag = diagnose_api_error("azure", api_err, model_or_deployment=resolved_dep, endpoint=base_endpoint)

                    if diag["kind"] in ["init_quota", "auth", "forbidden", "not_found", "bad_request"]:
                        msg = format_api_error_message(diag)
                        shared_context["fatal_error"] = msg
                        raise Exception(msg)
                    elif diag["kind"] == "rate_limit":
                        m = re.search(r'retry after (\d+)', err_lower)
                        if not m:
                            m = re.search(r'try again in (\d+)\s*seconds', err_lower)
                        if m:
                            wait_sec = float(m.group(1))
                            if wait_sec > 30.0:
                                msg = format_api_error_message(diag)
                                shared_context["fatal_error"] = msg
                                raise Exception(msg)
                            required_sleep = wait_sec + 2.0
                        else:
                            if attempt >= 3:
                                msg = format_api_error_message(diag)
                                shared_context["fatal_error"] = msg
                                raise Exception(msg)
                            required_sleep = 5.0 + random.uniform(1.0, 3.0)

                if success:
                    break
                if shared_context["fatal_error"]:
                    raise Exception(shared_context["fatal_error"])
                if ui.is_cancelled():
                    return

                if required_sleep > 0:
                    sleep_time = required_sleep
                else:
                    sleep_time = min(45.0, 3.0 * (2 ** attempt)) + random.uniform(1.0, 2.0)

                wait_step = 0.5
                while sleep_time > 0:
                    if shared_context["fatal_error"]:
                        raise Exception(shared_context["fatal_error"])
                    if ui.is_cancelled():
                        return
                    time.sleep(min(wait_step, sleep_time))
                    sleep_time -= wait_step

            _parse_and_save_page_result(
                extracted_text, success, last_error,
                is_table_format, crop_regions, cropped_info,
                page_num, total_p, base, digits, save_dir, out_format
            )
            gc.collect()

        except Exception as e:
            if shared_context["fatal_error"]:
                raise Exception(shared_context["fatal_error"])
            raise Exception(f"ページ {page_num+1} の処理中にエラーが発生しました: {e}")

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(process_single_page_task, task) for task in page_tasks]

        for future in concurrent.futures.as_completed(futures):
            if ui.is_cancelled() or shared_context["fatal_error"]:
                executor.shutdown(wait=False, cancel_futures=True)
                if shared_context["fatal_error"]:
                    raise Exception(shared_context["fatal_error"])
                return

            try:
                future.result()
            except Exception as e:
                executor.shutdown(wait=False, cancel_futures=True)
                raise e

            with progress_lock:
                completed_pages += 1
                ui.update_overall(completed_pages, total_tasks, f"全体の進捗 ( {completed_pages} / {total_tasks} ページ完了 )")

    ui.set_determinate(total_tasks, total_tasks, "すべての処理が完了しました")


# ==============================
# 各AIエンジン 最新モデル一覧取得機能
# ==============================
def fetch_gemini_models(api_key):
    """Gemini API から最新の利用可能な生成モデル一覧を取得し、新しいバージョン順にソートして返す"""
    if not api_key or not api_key.strip():
        raise Exception("APIキーが入力されていません。")

    genai.configure(api_key=api_key.strip())
    new_models = []
    exclude_keywords = ["tts", "audio", "image", "vision", "embedding", "aqa", "learnlm"]

    for m in genai.list_models():
        methods = getattr(m, "supported_generation_methods", []) or []
        if "generateContent" in methods and "gemini" in m.name.lower():
            model_id = m.name.replace("models/", "")
            if any(k in model_id.lower() for k in exclude_keywords):
                continue
            display_name = getattr(m, "display_name", model_id) or model_id
            new_models.append((f"{display_name} ({model_id})", model_id))

    def _gemini_sort_key(item):
        mid = item[1].lower()
        ver_match = re.search(r'gemini-(\d+(?:\.\d+)?)', mid)
        ver_num = float(ver_match.group(1)) if ver_match else 0.0
        is_preview = 1 if ("preview" in mid or "exp" in mid) else 0
        is_flash = 1 if "flash" in mid else 0
        return (-ver_num, is_preview, -is_flash, mid)

    new_models.sort(key=_gemini_sort_key)
    return new_models


def get_recommended_gemini_settings(model_id, is_free=True):
    """Gemini モデルとプラン（無料枠/課金枠）に応じた推奨設定を返す"""
    mid = (model_id or "").lower()
    if any(k in mid for k in ["carbon", "pro"]):
        rpm = 1 if is_free else 100
        threads = 1 if is_free else 3
    elif "neon" in mid:
        rpm = 2 if is_free else 150
        threads = 1 if is_free else 5
    elif "argon" in mid:
        rpm = 5 if is_free else 200
        threads = 1 if is_free else 5
    else:
        rpm = 12 if is_free else 300
        threads = 1 if is_free else 5
    return {
        "rpm": rpm,
        "threads": threads,
        "temperature": 0.0,
        "max_tokens": 8192,
        "disable_safety": True
    }


def get_recommended_azure_settings(deployment_name, detected_api_version=None):
    """Azure OpenAI デプロイ名（モデル種別）に応じた推奨設定を返す"""
    dep = (deployment_name or "").lower()
    api_ver = (detected_api_version or "").strip() or DEFAULT_AZURE_API_VERSION
    if any(k in dep for k in ["mini", "nano"]) and not any(dep.startswith(p) for p in ["o1", "o3", "o4"]):
        rpm = 120
        threads = 5
        max_tokens = 8192
    elif any(dep.startswith(p) or f"-{p}" in dep for p in ["gpt-6", "o4", "o3", "o1"]) or "pro" in dep:
        rpm = 30
        threads = 2
        max_tokens = 16384
    else:
        rpm = 60
        threads = 3
        max_tokens = 8192
    return {
        "api_version": api_ver,
        "rpm": rpm,
        "threads": threads,
        "temperature": 0.0,
        "max_tokens": max_tokens
    }


def fetch_azure_models_with_meta(api_key, endpoint, api_version=DEFAULT_AZURE_API_VERSION):
    """
    Azure OpenAI Service / Azure AI Foundry から利用可能な最新モデル一覧を取得し、
    同時に通信に成功した最新の推奨APIバージョンも判定して (model_ids, best_api_version) を返す。
    """
    base_endpoint, _, resolved_ver = parse_azure_endpoint_info(endpoint, "", api_version)
    if not base_endpoint:
        raise Exception("エンドポイントURLが正しく入力されていません。")
    if not api_key or not api_key.strip():
        raise Exception("APIキーが入力されていません。")

    api_key = api_key.strip()
    # 新しい推奨バージョンから順に試行して、利用可能な最新APIバージョンも同時に特定する
    versions_to_try = []
    for v in list(AZURE_API_VERSION_CANDIDATES) + [resolved_ver]:
        if v and v not in versions_to_try:
            versions_to_try.append(v)

    candidate_requests = [
        (f"{base_endpoint.rstrip('/')}/openai/models?api-version={urllib.parse.quote(v)}", v)
        for v in versions_to_try
    ]
    candidate_requests.append((f"{base_endpoint.rstrip('/')}/openai/v1/models", resolved_ver or DEFAULT_AZURE_API_VERSION))

    raw_data = None
    best_api_version = resolved_ver or DEFAULT_AZURE_API_VERSION
    last_err = None
    for url, ver_label in candidate_requests:
        req = urllib.request.Request(
            url,
            headers={
                "Content-Type": "application/json",
                "api-key": api_key
            },
            method="GET"
        )
        try:
            with urllib.request.urlopen(req, timeout=30.0) as resp:
                raw_data = json.loads(resp.read().decode("utf-8"))
                best_api_version = ver_label
                break
        except urllib.error.HTTPError as http_err:
            err_body = ""
            try:
                err_body = http_err.read().decode("utf-8", errors="replace")
            except Exception:
                pass
            last_err = Exception(f"HTTP {http_err.code}: {err_body or http_err.reason}")
            if http_err.code in (401, 403):
                raise last_err
        except Exception as e:
            last_err = e

    if raw_data is None:
        raise last_err if last_err else Exception("モデル一覧の取得に失敗しました。")

    items = raw_data.get("data", []) if isinstance(raw_data, dict) else (raw_data if isinstance(raw_data, list) else [])
    exclude_keywords = [
        "embedding", "whisper", "tts", "dall-e", "sora", "audio",
        "realtime", "transcribe", "moderation", "babbage", "davinci",
        "ada", "curie", "text-", "codex", "instruct"
    ]

    seen = set()
    model_ids = []
    for item in items:
        if not isinstance(item, dict):
            continue
        mid = str(item.get("id") or item.get("model") or "").strip()
        if not mid or mid in seen:
            continue
        status = str(item.get("status", "")).lower()
        if status and status not in ("succeeded", "active", "available"):
            continue
        caps = item.get("capabilities")
        if isinstance(caps, dict) and caps.get("chat_completion") is False:
            continue
        mid_lower = mid.lower()
        if any(k in mid_lower for k in exclude_keywords):
            continue
        seen.add(mid)
        model_ids.append(mid)

    def _azure_sort_key(mid):
        m = mid.lower()
        ver_m = re.match(r'^gpt-(\d+(?:\.\d+)?)', m)
        if ver_m:
            ver_val = float(ver_m.group(1))
            if ver_val >= 6.0:
                rank = 0
            elif ver_val >= 5.0:
                rank = 1
            elif ver_val >= 4.5:
                rank = 3
            elif ver_val >= 4.1:
                rank = 4
            else:
                rank = 6
            sub_ver = -ver_val
        elif m.startswith("o4"):
            rank, sub_ver = 2, 0.0
        elif m.startswith("gpt-4o"):
            rank, sub_ver = 5, 0.0
        elif m.startswith("o3"):
            rank, sub_ver = 7, 0.0
        elif m.startswith("o1"):
            rank, sub_ver = 8, 0.0
        else:
            rank, sub_ver = 9, 0.0
        has_date = 1 if re.search(r'-20\d{2}-', m) else 0
        return (rank, sub_ver, has_date, m)

    model_ids.sort(key=_azure_sort_key)
    return model_ids, best_api_version


def fetch_azure_models(api_key, endpoint, api_version=DEFAULT_AZURE_API_VERSION):
    """後方互換ラッパー: モデルIDのリストのみを返す"""
    model_ids, _ = fetch_azure_models_with_meta(api_key, endpoint, api_version)
    return model_ids


def filter_mainstream_gemini_models(models_list):
    """
    Gemini のモデル一覧 [(display_name, model_id), ...] から
    現在（最新世代）の主力モデルのみを動的に抽出する。
    ・2026年11月現在: Gemini 4系 (Hydrogen/Carbon/Argon/Neon等) ＋ 3.5以上の最新Flash/Flash-Lite系
    ・旧世代 (2.5 / 2.0 / 1.5等)・実験版・番号/日付付きスナップショットを除外
    """
    if not models_list:
        return []

    def _extract_ver(mid):
        m = re.search(r'gemini-(\d+(?:\.\d+)?)', mid.lower())
        return float(m.group(1)) if m else 0.0

    # 実験版・固定スナップショットを除外したクリーン候補を作成
    clean_candidates = []
    for item in models_list:
        mid = item[1].lower()
        if any(k in mid for k in ["-exp", "experimental", "thinking-exp", "-latest", "tuning", "robotics"]):
            continue
        if re.search(r'-(00\d|\d{2}-\d{2}|20\d{2})', mid):
            continue
        clean_candidates.append(item)

    if not clean_candidates:
        return models_list[:8]

    max_ver = max((_extract_ver(item[1]) for item in clean_candidates), default=0.0)

    mainstream = []
    for item in clean_candidates:
        mid = item[1].lower()
        ver = _extract_ver(mid)
        if max_ver >= 4.0:
            # Gemini 4系以上が主力。Flash/Flash-Lite系のみ 3.5以上を軽量主力として残す
            if ver >= 4.0 or (ver >= 3.5 and "flash" in mid):
                mainstream.append(item)
        elif max_ver >= 3.0:
            # 4系がまだ存在しない環境では 3.x系のみを主力とする（2.5系以下の旧世代は除外）
            if ver >= 3.0:
                mainstream.append(item)
        else:
            if ver >= max_ver - 0.5:
                mainstream.append(item)

    return mainstream if mainstream else clean_candidates[:8]


def filter_mainstream_azure_models(models_list):
    """
    Azure OpenAI のモデル一覧 [model_id, ...] から現在（最新世代）の主力モデルを抽出する。
    ・2026年11月現在: GPT-6系 / GPT-5系 (gpt-5, gpt-5.1, gpt-5.2, gpt-5-mini/nano/chat/pro等) / o4系を主力とする
    ・旧世代 (gpt-4o, gpt-4.1, gpt-4, o3, o1, gpt-35等) は、最新世代が存在する環境では除外する
    ・日付付きモデルIDしか返らない環境では各主力ファミリーの最新1件ずつを自動選出する
    """
    if not models_list:
        return []

    current_prefixes = ("gpt-6", "gpt-5", "o4-mini", "o4")
    legacy_prefixes = ("gpt-4.5", "gpt-4.1", "gpt-4o", "o3-mini", "o3")
    date_pattern = re.compile(r'-(20\d{2}-\d{2}-\d{2}|\d{4})$')
    exclude_sub = ("preview-20", "search", "transcribe", "diarize", "tts", "audio", "codex")

    def _pick_by_prefixes(prefixes):
        clean = []
        for mid in models_list:
            m = mid.lower()
            if not any(m == p or m.startswith(p + "-") or m.startswith(p + ".") for p in prefixes):
                continue
            if date_pattern.search(m):
                continue
            if any(k in m for k in exclude_sub):
                continue
            clean.append(mid)
        if clean:
            return clean

        # 日付付きIDしか存在しない場合は各ファミリーの最新1件を選出
        family_latest = {}
        for mid in models_list:
            m = mid.lower()
            if not any(m == p or m.startswith(p + "-") or m.startswith(p + ".") for p in prefixes):
                continue
            if any(k in m for k in exclude_sub):
                continue
            fam = date_pattern.sub("", m)
            if fam not in family_latest or mid > family_latest[fam]:
                family_latest[fam] = mid
        if family_latest:
            return [mid for mid in models_list if mid in family_latest.values()]
        return []

    # 第1優先: 現行主力世代 (GPT-6 / GPT-5 / o4)
    current_main = _pick_by_prefixes(current_prefixes)
    if current_main:
        return current_main

    # フォールバック: 現行世代がまだ展開されていないリージョンでは前世代 (GPT-4.1 / 4o / o3)
    legacy_main = _pick_by_prefixes(legacy_prefixes)
    if legacy_main:
        return legacy_main

    return models_list[:8]
