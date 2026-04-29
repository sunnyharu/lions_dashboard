"""
사내 Presto DB에서 온라인 매출(pay_amt)을 조회해
Google Sheets '일별매출' 탭의 ON매출 컬럼에 적재
"""
import json
import os
from datetime import datetime, timedelta

import prestodb
import pandas as pd
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
import gspread

load_dotenv()

# ── Presto 연결 설정 ────────────────────────────────────
PRESTO_HOST = "kakaoent-presto-adhoc.kakaoent.io"
PRESTO_PORT = 8443
PRESTO_USER = "journi-y222"
PRESTO_PASS = os.environ.get("PRESTO_PASS", "")

# ── Google Sheets 설정 ───────────────────────────────────
SPREADSHEET_ID    = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"
SHEET_NAME        = "일별매출"
GOOGLE_CREDS_ENV  = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE = "google_credentials.json"

# ── 조회 기준: 어제 ─────────────────────────────────────
yesterday = datetime.today() - timedelta(days=1)
DATE_STR  = yesterday.strftime("%Y.%m.%d")   # 시트 날짜 포맷: "2026.04.28"
DATE_SQL  = yesterday.strftime("%Y-%m-%d")   # SQL 날짜 포맷: "2026-04-28"

SQL = f"""
SELECT
    completed_dt,
    SUM(COALESCE(line_krw_amount, 0))
    + SUM(IF(delivery_rk = 1, COALESCE(krw_delivery_fee, 0), 0))
    - SUM(COALESCE(real_discount_amount, 0)) AS pay_amt
FROM data_analysis.v_berriz_commerce_mart_daily_order
WHERE partner_id = 6
  AND completed_dt = DATE '{DATE_SQL}'
GROUP BY 1
"""


def fetch_online_sales() -> int | None:
    conn = prestodb.dbapi.connect(
        host        = PRESTO_HOST,
        port        = PRESTO_PORT,
        user        = PRESTO_USER,
        catalog     = "hadoop_kent",
        schema      = "data_analysis",
        http_scheme = "https",
        auth        = prestodb.auth.BasicAuthentication(PRESTO_USER, PRESTO_PASS),
    )
    cursor = conn.cursor()
    cursor.execute(SQL)
    rows = cursor.fetchall()
    conn.close()

    if not rows:
        print(f"ON매출 없음: {DATE_SQL}")
        return None

    pay_amt = rows[0][1]
    print(f"ON매출({DATE_SQL}): {pay_amt:,.0f}")
    return int(pay_amt) if pay_amt else 0


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


def update_sheet(pay_amt: int):
    client = get_gspread_client()
    sh     = client.open_by_key(SPREADSHEET_ID)
    ws     = sh.worksheet(SHEET_NAME)

    all_values = ws.get_all_values()
    if not all_values:
        print("시트가 비어 있음")
        return

    header = all_values[0]

    # ON매출 컬럼이 없으면 추가
    if "ON매출" not in header:
        on_col = len(header) + 1
        ws.update_cell(1, on_col, "ON매출")
        header.append("ON매출")
        print(f"ON매출 헤더 추가 (열 {on_col})")
    else:
        on_col = header.index("ON매출") + 1

    # 날짜 컬럼 인덱스
    date_col = header.index("날짜") if "날짜" in header else 0

    # 해당 날짜 행 찾기
    target_row = None
    for i, row in enumerate(all_values[1:], start=2):
        if row and row[date_col] == DATE_STR:
            target_row = i
            break

    if target_row:
        ws.update_cell(target_row, on_col, pay_amt)
        print(f"ON매출 업데이트: {DATE_STR} → {pay_amt:,} (행 {target_row})")
    else:
        # 해당 날짜 행이 없으면 새 행 추가
        new_row = [""] * (on_col - 1)
        new_row[date_col] = DATE_STR
        new_row[on_col - 1] = pay_amt
        ws.append_row(new_row)
        print(f"새 행 추가: {DATE_STR}, ON매출={pay_amt:,}")


def main():
    if not PRESTO_PASS:
        raise ValueError("PRESTO_PASS 환경변수가 없습니다")

    pay_amt = fetch_online_sales()
    if pay_amt is not None:
        update_sheet(pay_amt)


if __name__ == "__main__":
    main()
