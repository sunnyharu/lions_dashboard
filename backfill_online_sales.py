"""
온라인 매출 백필 스크립트
Presto에서 지정 기간 pay_amt 조회 → 일별매출 시트 ON매출 컬럼 일괄 업데이트
"""
import json
import os
from datetime import datetime, timedelta, date

import prestodb
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

# ── 백필 기간 ────────────────────────────────────────────
_start_env     = os.environ.get("BACKFILL_START", "2026-02-24")
_end_env       = os.environ.get("BACKFILL_END",   (datetime.today() - timedelta(days=1)).strftime("%Y-%m-%d"))
BACKFILL_START = date.fromisoformat(_start_env)
BACKFILL_END   = date.fromisoformat(_end_env)

SQL = f"""
SELECT
    completed_dt,
    SUM(COALESCE(line_krw_amount, 0))
    + SUM(IF(delivery_rk = 1, COALESCE(krw_delivery_fee, 0), 0))
    - SUM(COALESCE(real_discount_amount, 0)) AS pay_amt
FROM data_analysis.v_berriz_commerce_mart_daily_order
WHERE partner_id = 6
  AND completed_dt BETWEEN DATE '{BACKFILL_START}' AND DATE '{BACKFILL_END}'
GROUP BY 1
ORDER BY 1
"""


def fetch_all_online_sales() -> dict:
    """기간 내 날짜별 pay_amt 반환 {date: amount}"""
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
    print(f"Presto 조회 중: {BACKFILL_START} ~ {BACKFILL_END}")
    cursor.execute(SQL)
    rows = cursor.fetchall()
    conn.close()

    result = {}
    for completed_dt, pay_amt in rows:
        # completed_dt는 date 객체 또는 문자열
        if hasattr(completed_dt, "strftime"):
            key = completed_dt.strftime("%Y.%m.%d")
        else:
            key = str(completed_dt).replace("-", ".")  # "2026.02.24"
            # "2026.2.24" → "2026.02.24" 정규화
            parts = key.split(".")
            key = f"{parts[0]}.{int(parts[1]):02d}.{int(parts[2]):02d}"
        result[key] = int(pay_amt) if pay_amt else 0

    print(f"조회된 날짜 수: {len(result)}")
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


def update_sheet(sales_by_date: dict):
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

    date_col_idx = header.index("날짜") if "날짜" in header else 0

    # 날짜 → 행번호 매핑
    date_to_row = {
        row[date_col_idx]: i + 2
        for i, row in enumerate(all_values[1:])
        if row and row[date_col_idx]
    }

    # 일괄 업데이트 (batch_update)
    updates = []
    missing_dates = []
    for date_str, amount in sorted(sales_by_date.items()):
        if date_str in date_to_row:
            row_num = date_to_row[date_str]
            cell = gspread.utils.rowcol_to_a1(row_num, on_col)
            updates.append({"range": cell, "values": [[amount]]})
        else:
            missing_dates.append((date_str, amount))

    if updates:
        ws.batch_update(updates)
        print(f"업데이트 완료: {len(updates)}일")

    # 시트에 없는 날짜는 새 행 추가
    if missing_dates:
        print(f"시트에 없는 날짜 {len(missing_dates)}건 → 새 행 추가")
        for date_str, amount in missing_dates:
            new_row = [""] * len(header)
            new_row[date_col_idx] = date_str
            new_row[on_col - 1]   = amount
            ws.append_row(new_row)
            print(f"  추가: {date_str} = {amount:,}")


def main():
    if not PRESTO_PASS:
        raise ValueError("PRESTO_PASS 환경변수가 없습니다")

    sales = fetch_all_online_sales()
    if sales:
        update_sheet(sales)
    else:
        print("조회된 매출 데이터 없음")


if __name__ == "__main__":
    main()
