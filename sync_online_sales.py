"""
data/online_sales.xlsx 읽어서 일별매출 시트 ON거래액 컬럼 업데이트
컬럼: completed_dt (date), pay_amt (float)
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


def load_excel() -> dict:
    """CSV → {날짜문자열: 금액} 딕셔너리 반환"""
    df = pd.read_csv(CSV_PATH)
    df.columns = df.columns.str.strip()

    result = {}
    for _, row in df.iterrows():
        dt  = pd.to_datetime(row["completed_dt"])
        amt = int(row["pay_amt"]) if pd.notna(row["pay_amt"]) else 0
        key = dt.strftime("%Y.%m.%d")   # 시트 포맷: "2026.02.24"
        result[key] = amt

    print(f"Excel 로드: {len(result)}일치")
    return result


def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    if GOOGLE_CREDS_ENV:
        info = json.loads(GOOGLE_CREDS_ENV)
        creds = Credentials.from_service_account_info(info, scopes=scopes)
    else:
        creds = Credentials.from_service_account_file(GOOGLE_CREDS_FILE, scopes=scopes)
    return gspread.authorize(creds)


def update_sheet(sales: dict):
    client = get_gspread_client()
    ws     = get_gspread_client().open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
    ws     = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)

    all_values = ws.get_all_values()
    if not all_values:
        print("시트가 비어 있음")
        return

    header = all_values[0]

    # ON거래액 컬럼 없으면 추가
    if "ON거래액" not in header:
        on_col = len(header) + 1
        ws.update_cell(1, on_col, "ON거래액")
        header.append("ON거래액")
        print(f"ON거래액 헤더 추가 (열 {on_col})")
    else:
        on_col = header.index("ON거래액") + 1

    date_col_idx = header.index("날짜") if "날짜" in header else 0

    def normalize_date(s):
        """'2026.4.1' → '2026.04.01' 정규화"""
        try:
            parts = s.split(".")
            return f"{parts[0]}.{int(parts[1]):02d}.{int(parts[2]):02d}"
        except Exception:
            return s

    # 날짜 → 행번호 매핑 (정규화된 날짜 기준)
    date_to_row = {
        normalize_date(row[date_col_idx]): i + 2
        for i, row in enumerate(all_values[1:])
        if row and row[date_col_idx]
    }

    # 기존 행 업데이트
    updates      = []
    missing      = []
    for date_str, amount in sorted(sales.items()):
        if date_str in date_to_row:
            cell = gspread.utils.rowcol_to_a1(date_to_row[date_str], on_col)
            updates.append({"range": cell, "values": [[amount]]})
        else:
            missing.append((date_str, amount))

    if updates:
        ws.batch_update(updates)
        print(f"업데이트 완료: {len(updates)}일")

    # 시트에 없는 날짜는 새 행 추가
    for date_str, amount in missing:
        new_row = [""] * len(header)
        new_row[date_col_idx] = date_str
        new_row[on_col - 1]   = amount
        ws.append_row(new_row)
        print(f"새 행 추가: {date_str} = {amount:,}")


def main():
    sales = load_excel()
    if sales:
        update_sheet(sales)


if __name__ == "__main__":
    main()
