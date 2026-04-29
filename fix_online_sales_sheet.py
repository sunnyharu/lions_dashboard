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

    date_col_idx = header.index("날짜")      if "날짜"      in header else 0
    off_col_idx  = header.index("일별거래액") if "일별거래액" in header else 1

    # ── 1. CSV 로드 ──
    df = pd.read_csv(CSV_PATH)
    df.columns = df.columns.str.strip()
    sales = {}
    for _, row in df.iterrows():
        dt  = pd.to_datetime(row["completed_dt"])
        amt = int(row["pay_amt"]) if pd.notna(row["pay_amt"]) else 0
        sales[dt.strftime("%Y.%m.%d")] = amt

    # ── 2. 유효한 행만 필터 + ON매출 병합 ──
    new_rows = [["날짜", "일별거래액", "ON매출"]]
    kept, removed = 0, 0
    for row in all_values[1:]:
        date_val = row[date_col_idx].strip() if date_col_idx < len(row) else ""
        off_val  = row[off_col_idx].strip()  if off_col_idx  < len(row) else ""

        # 일별거래액이 없는 행은 제거
        if not off_val:
            removed += 1
            continue

        norm_date = normalize_date(date_val)
        on_amt    = sales.get(norm_date, "")
        new_rows.append([date_val, off_val, on_amt])
        kept += 1

    print(f"제거할 행: {removed}개 / 유지할 행: {kept}개")

    # ── 3. 시트 전체를 한 번에 교체 (API 호출 최소화) ──
    ws.clear()
    ws.update(new_rows, value_input_option="USER_ENTERED")
    print(f"시트 재작성 완료: {kept}행, ON매출 {sum(1 for r in new_rows[1:] if r[2] != '')}일 채움")


if __name__ == "__main__":
    main()
