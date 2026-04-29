"""
일별매출 시트 정리:
1. ON매출만 있고 날짜(일별거래액)가 없는 잘못 추가된 행 삭제
2. 날짜 정규화 후 ON매출 재동기화
"""
import json
import os

import pandas as pd
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
import gspread

load_dotenv()

SPREADSHEET_ID    = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"
SHEET_NAME        = "일별매출"
GOOGLE_CREDS_ENV  = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE = "google_credentials.json"
CSV_PATH          = "data/online_sales.csv"


def normalize_date(s):
    try:
        parts = s.split(".")
        return f"{parts[0]}.{int(parts[1]):02d}.{int(parts[2]):02d}"
    except Exception:
        return s


def get_ws():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    if GOOGLE_CREDS_ENV:
        info = json.loads(GOOGLE_CREDS_ENV)
        creds = Credentials.from_service_account_info(info, scopes=scopes)
    else:
        creds = Credentials.from_service_account_file(GOOGLE_CREDS_FILE, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)


def main():
    ws         = get_ws()
    all_values = ws.get_all_values()
    header     = all_values[0]

    date_col_idx    = header.index("날짜")      if "날짜"   in header else 0
    off_col_idx     = header.index("일별거래액") if "일별거래액" in header else 1
    on_col          = header.index("ON매출") + 1 if "ON매출" in header else None

    # ── 1. 잘못 추가된 행 삭제 (날짜는 있는데 일별거래액이 없는 행) ──
    rows_to_delete = []
    for i, row in enumerate(all_values[1:], start=2):
        has_date     = row[date_col_idx].strip() if date_col_idx < len(row) else ""
        has_off      = row[off_col_idx].strip()  if off_col_idx  < len(row) else ""
        if has_date and not has_off:
            rows_to_delete.append(i)

    # 뒤에서부터 삭제해야 행 번호가 안 밀림
    for row_num in reversed(rows_to_delete):
        ws.delete_rows(row_num)
        print(f"삭제: 행 {row_num}")

    print(f"잘못된 행 {len(rows_to_delete)}개 삭제 완료")

    # ── 2. ON매출 컬럼 확보 ──
    all_values = ws.get_all_values()
    header     = all_values[0]

    if "ON매출" not in header:
        on_col = len(header) + 1
        ws.update_cell(1, on_col, "ON매출")
        header.append("ON매출")
    else:
        on_col = header.index("ON매출") + 1

    # ── 3. CSV 로드 ──
    df = pd.read_csv(CSV_PATH)
    df.columns = df.columns.str.strip()
    sales = {}
    for _, row in df.iterrows():
        dt  = pd.to_datetime(row["completed_dt"])
        amt = int(row["pay_amt"]) if pd.notna(row["pay_amt"]) else 0
        sales[dt.strftime("%Y.%m.%d")] = amt

    # ── 4. 날짜 매핑 (정규화) ──
    date_to_row = {
        normalize_date(row[date_col_idx]): i + 2
        for i, row in enumerate(all_values[1:])
        if row and row[date_col_idx]
    }

    updates = []
    for date_str, amount in sorted(sales.items()):
        if date_str in date_to_row:
            cell = gspread.utils.rowcol_to_a1(date_to_row[date_str], on_col)
            updates.append({"range": cell, "values": [[amount]]})

    if updates:
        ws.batch_update(updates)
        print(f"ON매출 업데이트 완료: {len(updates)}일")


if __name__ == "__main__":
    main()
