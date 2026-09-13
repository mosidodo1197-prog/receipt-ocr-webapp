# 영수증 OCR 자동 정리 웹앱 (Streamlit)
# - 영수증 이미지를 업로드하면 Upstage OCR로 텍스트를 추출하고,
#   Upstage LLM(JSON Schema)으로 가게이름/구매일시/물품/총구매금액을 구조화하여 표로 보여줌
# - 결과는 엑셀(.xlsx)로 다운로드 가능 (파일명·가게이름 셀 병합 포함)

import io
import json
import os
import re
import time

import openpyxl
import pandas as pd
import requests
import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI
from openpyxl.styles import Alignment

load_dotenv()  # 로컬 개발 시 .env 파일을 읽음 (배포 환경에서는 st.secrets 사용)


def get_config(key, default=None):
    """배포 환경(st.secrets)과 로컬 환경(.env) 양쪽을 모두 지원"""
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return os.environ.get(key, default)


UPSTAGE_API_KEY = get_config("UPSTAGE_API_KEY")
UPSTAGE_BASE_URL = get_config("UPSTAGE_BASE_URL", "https://api.upstage.ai/v1")

OCR_URL = "https://api.upstage.ai/v1/document-digitization"

RECEIPT_SCHEMA = {
    "type": "object",
    "properties": {
        "store_name": {"type": "string", "description": "가게(매장) 이름. 반드시 원문 텍스트에 있는 글자 그대로 옮겨 적을 것"},
        "purchase_datetime": {"type": "string", "description": "구매 일시"},
        "items": {
            "type": "array",
            "description": "구매한 물품 목록",
            "items": {
                "type": "object",
                "properties": {
                    "item_name": {"type": "string", "description": "수량, 단가 표기를 제외한 순수 물품명"},
                    "quantity": {"type": "number", "description": "구매 수량"},
                    "unit_price": {"type": "number", "description": "물품 1개당 단가"},
                },
                "required": ["item_name", "quantity", "unit_price"],
                "additionalProperties": False,
            },
        },
        "total_amount": {"type": "number", "description": "총 구매금액"},
    },
    "required": ["store_name", "purchase_datetime", "items", "total_amount"],
    "additionalProperties": False,
}

RESULT_COLUMNS = ["파일명", "가게이름", "구매일시", "물품명", "수량", "단가", "금액", "총 구매금액"]


def call_with_retry(fn, max_retries=5, base_delay=3):
    """429(Too Many Requests)를 만나면 잠시 대기 후 자동 재시도"""
    for attempt in range(max_retries):
        try:
            return fn()
        except Exception as e:
            status_code = getattr(e, "status_code", None)
            response = getattr(e, "response", None)
            if status_code is None and response is not None:
                status_code = getattr(response, "status_code", None)

            is_last_attempt = attempt == max_retries - 1
            if status_code == 429 and not is_last_attempt:
                retry_after = None
                if response is not None:
                    retry_after = response.headers.get("Retry-After")
                delay = float(retry_after) if retry_after else base_delay * (2 ** attempt)
                time.sleep(delay)
                continue
            raise


def fix_store_name(store_name, ocr_text):
    """LLM이 가끔 가게이름을 엉뚱하게 지어내는 경우(예: 이상한 한글 조합)를 보정.
    OCR 원문에 실제로 등장하는 문자열인지 확인하고, 아니면 원문 첫 줄(대개 가게이름)로 대체"""
    normalized_text = ocr_text.replace(" ", "")
    normalized_name = (store_name or "").replace(" ", "")
    if normalized_name and normalized_name in normalized_text:
        return store_name

    first_line = ocr_text.strip().splitlines()[0].strip() if ocr_text.strip() else store_name
    return first_line


ITEM_LINE_PATTERN = re.compile(r"^(.*\S)\s+(\d+)\s+([\d,]+)\s+([\d,]+)\s*$")


def parse_ocr_item_candidates(ocr_text):
    """OCR 텍스트에서 '물품명 수량 단가 금액' 형태의 줄을 찾아 후보 목록으로 반환"""
    candidates = []
    for line in ocr_text.splitlines():
        match = ITEM_LINE_PATTERN.match(line.strip())
        if not match:
            continue
        name, qty, price, amount = match.groups()
        try:
            candidates.append(
                {
                    "name": name.strip(),
                    "quantity": int(qty),
                    "unit_price": int(price.replace(",", "")),
                }
            )
        except ValueError:
            continue
    return candidates


def fix_item_name(item_name, quantity, unit_price, ocr_text, candidates):
    """LLM이 물품명을 엉뚱하게 지어내는 경우를 보정.
    OCR 원문에 없는 이름이면, 같은 수량·단가를 가진 원문 줄에서 실제 물품명을 찾아 대체"""
    normalized_text = ocr_text.replace(" ", "")
    normalized_name = (item_name or "").replace(" ", "")
    if normalized_name and normalized_name in normalized_text:
        return item_name

    for candidate in candidates:
        if candidate["quantity"] == quantity and candidate["unit_price"] == round(unit_price):
            return candidate["name"]
    return item_name  # 매칭되는 원문 줄을 못 찾으면 원래 값 유지


def extract_receipt_rows(client, file_bytes, filename):
    """영수증 이미지(메모리 바이트) 1개를 OCR + LLM으로 처리하여 물품별 행 리스트를 반환"""
    headers = {"Authorization": f"Bearer {UPSTAGE_API_KEY}"}

    def do_ocr():
        resp = requests.post(
            OCR_URL,
            headers=headers,
            files={"document": (filename, io.BytesIO(file_bytes))},
            data={"model": "ocr"},
            timeout=60,
        )
        resp.raise_for_status()
        return resp

    ocr_response = call_with_retry(do_ocr)
    ocr_text = ocr_response.json()["text"]

    def do_llm():
        return client.chat.completions.create(
            model="solar-mini",
            timeout=60,
            temperature=0,
            messages=[
                {
                    "role": "system",
                    "content": "너는 영수증 OCR 텍스트에서 정보를 정확하게 추출하는 도우미야. "
                    "원문에 있는 글자를 절대 지어내지 말고 그대로 옮겨 적어. "
                    "물품명(item_name)에는 수량이나 가격 표기를 포함하지 말고 순수 상품명만 넣어.",
                },
                {"role": "user", "content": ocr_text},
            ],
            response_format={
                "type": "json_schema",
                "json_schema": {"name": "receipt_info", "schema": RECEIPT_SCHEMA, "strict": True},
            },
        )

    llm_response = call_with_retry(do_llm)
    info = json.loads(llm_response.choices[0].message.content)
    store_name = fix_store_name(info["store_name"], ocr_text)
    item_candidates = parse_ocr_item_candidates(ocr_text)

    return [
        {
            "파일명": filename,
            "가게이름": store_name,
            "구매일시": info["purchase_datetime"],
            "물품명": fix_item_name(item["item_name"], item["quantity"], item["unit_price"], ocr_text, item_candidates),
            "수량": item["quantity"],
            "단가": item["unit_price"],
            "금액": round(item["quantity"] * item["unit_price"], 2),
            "총 구매금액": info["total_amount"],
        }
        for item in info["items"]
    ]


def merge_by_group(ws, group_col_idx, target_col_idxs, n_rows, header_row=1):
    """group_col_idx 값이 연속으로 같은 구간을 찾아 target_col_idxs 열들을 함께 병합"""
    start_row = header_row + 1
    last_row = header_row + n_rows
    merge_start = start_row
    prev_key = ws.cell(row=start_row, column=group_col_idx).value
    for row in range(start_row + 1, last_row + 2):
        current_key = ws.cell(row=row, column=group_col_idx).value if row <= last_row else object()
        if current_key != prev_key:
            if row - 1 > merge_start:
                for col_idx in target_col_idxs:
                    ws.merge_cells(start_row=merge_start, start_column=col_idx, end_row=row - 1, end_column=col_idx)
                    ws.cell(row=merge_start, column=col_idx).alignment = Alignment(vertical="center")
            merge_start = row
            prev_key = current_key


def build_merged_excel_bytes(df):
    """DataFrame을 파일명/가게이름 셀 병합이 적용된 엑셀 바이트로 변환 (디스크에 저장하지 않음)"""
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)

    wb = openpyxl.load_workbook(buffer)
    ws = wb.active
    merge_by_group(ws, group_col_idx=1, target_col_idxs=[1, 2], n_rows=len(df))

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


# ==============================
# Streamlit UI
# ==============================
st.set_page_config(page_title="영수증 OCR 자동 정리", page_icon="🧾")
st.title("🧾 영수증 OCR 자동 정리")
st.caption("영수증 이미지를 업로드하면 Upstage OCR + LLM으로 항목을 자동 추출해 표로 정리합니다.")

if not UPSTAGE_API_KEY:
    st.error("UPSTAGE_API_KEY가 설정되어 있지 않습니다. .env 파일 또는 배포 환경의 Secrets를 확인하세요.")
    st.stop()

uploaded_files = st.file_uploader(
    "영수증 이미지를 업로드하세요 (여러 장 가능)",
    type=["jpg", "jpeg", "png", "webp", "bmp"],
    accept_multiple_files=True,
)

if st.button("분석하기", type="primary", disabled=not uploaded_files):
    client = OpenAI(api_key=UPSTAGE_API_KEY, base_url=UPSTAGE_BASE_URL)

    all_rows = []
    progress = st.progress(0.0)
    status_area = st.empty()
    for i, uploaded_file in enumerate(uploaded_files):
        status_area.info(f"({i + 1}/{len(uploaded_files)}) {uploaded_file.name} 처리 중...")
        try:
            rows = extract_receipt_rows(client, uploaded_file.getvalue(), uploaded_file.name)
            all_rows.extend(rows)
        except Exception as e:
            st.warning(f"{uploaded_file.name} 처리 실패: {e}")
        progress.progress((i + 1) / len(uploaded_files))
        if i < len(uploaded_files) - 1:
            time.sleep(1)  # 연속 호출로 인한 429(Too Many Requests) 예방
    status_area.empty()

    if all_rows:
        st.session_state["result_df"] = pd.DataFrame(all_rows, columns=RESULT_COLUMNS)
        st.success(f"영수증 {len(uploaded_files)}개 처리 완료!")

if "result_df" in st.session_state:
    result_df = st.session_state["result_df"]
    st.dataframe(result_df, use_container_width=True)

    excel_bytes = build_merged_excel_bytes(result_df)
    st.download_button(
        label="엑셀로 다운로드",
        data=excel_bytes,
        file_name="receipt_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
