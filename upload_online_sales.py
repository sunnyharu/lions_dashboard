"""
온라인 판매 CSV → 상품별매출(on) 시트 업로드
- 기존 데이터 전체 삭제 후 재적재 (헤더 유지)
- 바코드는 텍스트 형식으로 강제 처리 (정밀도 손실 방지)
"""
import csv
import json
import os
import sys
import time

import gspread
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials

load_dotenv()

SPREADSHEET_ID    = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"
SHEET_NAME        = "상품별매출(on)"
GOOGLE_CREDS_ENV  = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE = "google_credentials.json"

# CSV 컬럼 → 시트 컬럼 매핑
# CSV:   판매일, 상품ID, 상품명, 바코드, skucode, 사이즈, 선수명, 판매가, 결제상품수, 상품결제금액
SHEET_HEADER = ["판매일자", "상품ID", "상품명", "바코드", "skucode", "사이즈", "선수명", "판매단가", "판매수량", "실판매금액"]
CSV_COLUMNS  = ["판매일",  "상품ID", "상품명", "바코드", "skucode", "사이즈", "선수명", "판매가",   "결제상품수", "상품결제금액"]

# 바코드에서 variant suffix 제거 (e.g. "8804775462061_KYW2" → "8804775462061")
def clean_barcode(v: str) -> str:
    v = v.strip()
    if "_" in v:
        v = v.split("_")[0]
    return v


def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    if GOOGLE_CREDS_ENV:
        creds = Credentials.from_service_account_info(json.loads(GOOGLE_CREDS_ENV), scopes=scopes)
    else:
        creds = Credentials.from_service_account_file(GOOGLE_CREDS_FILE, scopes=scopes)
    return gspread.authorize(creds)


def read_csv(path: str) -> list:
    rows = []
    with open(path, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for raw in reader:
            row = []
            for col in CSV_COLUMNS:
                v = str(raw.get(col, "") or "").strip().strip('"')
                if col == "바코드":
                    v = clean_barcode(v)
                row.append(v)
            rows.append(row)
    print(f"CSV 읽기 완료: {len(rows)}행")
    return rows


def upload(csv_path: str):
    rows = read_csv(csv_path)
    if not rows:
        print("업로드할 데이터 없음")
        return

    client = get_gspread_client()
    sh = client.open_by_key(SPREADSHEET_ID)

    try:
        ws = sh.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=50000, cols=len(SHEET_HEADER))
        print(f"시트 '{SHEET_NAME}' 생성")

    # 헤더 아래 데이터 전체 삭제
    print("기존 데이터 삭제 중...")
    all_vals = ws.get_all_values()
    if len(all_vals) > 1:
        # 헤더 제외한 나머지 행 범위 clear
        last_row = len(all_vals)
        ws.batch_clear([f"A2:{chr(64 + len(SHEET_HEADER))}{last_row}"])
        print(f"  {last_row - 1}행 삭제 완료")

    # 헤더 확인 / 세팅
    current_header = ws.row_values(1) if all_vals else []
    if current_header != SHEET_HEADER:
        ws.update([SHEET_HEADER], "A1")
        print("헤더 업데이트 완료")

    # 바코드 컬럼 인덱스
    barcode_col_idx = SHEET_HEADER.index("바코드")  # 3

    # 바코드를 텍스트로 강제하기 위해 앞에 ' 붙임 (USER_ENTERED 모드에서 텍스트 처리)
    upload_rows = []
    for row in rows:
        r = row[:]
        if r[barcode_col_idx]:
            r[barcode_col_idx] = "'" + r[barcode_col_idx]
        upload_rows.append(r)

    # 500행씩 배치 업로드
    BATCH = 500
    total = len(upload_rows)
    for start in range(0, total, BATCH):
        chunk = upload_rows[start:start + BATCH]
        ws.append_rows(chunk, value_input_option="USER_ENTERED")
        print(f"  {min(start + BATCH, total)}/{total} 업로드 완료")
        if start + BATCH < total:
            time.sleep(1)

    print(f"✅ 업로드 완료: {total}행 → '{SHEET_NAME}' 시트")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("사용법: python upload_online_sales.py <CSV파일경로>")
        print("예시: python upload_online_sales.py ~/Desktop/2026-06-17-12:33.csv")
        sys.exit(1)
    upload(sys.argv[1])
