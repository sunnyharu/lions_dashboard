"""
Google Sheets → Presto 적재
- hadoop_kent.data_analysis.lions_daily_sales   (일별매출)
- hadoop_kent.data_analysis.lions_product_sales (상품별매출)

사내망에서만 실행 가능. 매일 출근 후 수동 or 작업스케줄러로 실행.
"""
import json
import os
import sys
from datetime import datetime

import prestodb
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
import gspread

load_dotenv()

# ── Presto 연결 ────────────────────────────────────────
PRESTO_HOST     = os.environ.get("PRESTO_HOST",     "kakaoent-presto-adhoc.kakaoent.io")
PRESTO_PORT     = int(os.environ.get("PRESTO_PORT", 8443))
PRESTO_USER     = os.environ.get("PRESTO_USER",     "")
PRESTO_PASSWORD = os.environ.get("PRESTO_PASSWORD", "")
PRESTO_CATALOG  = os.environ.get("PRESTO_CATALOG",  "hadoop_kent")
PRESTO_SCHEMA   = os.environ.get("PRESTO_SCHEMA",   "data_analysis")

# ── Google Sheets ──────────────────────────────────────
SPREADSHEET_ID   = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"
GOOGLE_CREDS_ENV  = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE = "google_credentials.json"

TABLE_DAILY   = f"{PRESTO_CATALOG}.{PRESTO_SCHEMA}.lions_daily_sales"
TABLE_PRODUCT = f"{PRESTO_CATALOG}.{PRESTO_SCHEMA}.lions_product_sales"


# ── 연결 ─────────────────────────────────────────────

def get_presto_conn():
    return prestodb.dbapi.connect(
        host        = PRESTO_HOST,
        port        = PRESTO_PORT,
        user        = PRESTO_USER,
        catalog     = PRESTO_CATALOG,
        schema      = PRESTO_SCHEMA,
        http_scheme = "https",
        auth        = prestodb.auth.BasicAuthentication(PRESTO_USER, PRESTO_PASSWORD),
    )


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


def execute(cur, sql, description=""):
    if description:
        print(f"  → {description}")
    cur.execute(sql)
    return cur


def to_int(v):
    try: return int(float(str(v).replace(",", "")))
    except: return 0


# ── 기존 날짜 조회 ─────────────────────────────────────

def get_existing_dates(cur, table: str) -> set:
    try:
        cur.execute(f"SELECT DISTINCT sale_date FROM {table}")
        return {row[0] for row in cur.fetchall()}
    except Exception:
        return set()


# ── 1. 일별매출 ───────────────────────────────────────

DDL_DAILY = f"""
CREATE TABLE IF NOT EXISTS {TABLE_DAILY} (
    sale_date    VARCHAR,
    off_amount   BIGINT,
    on_amount    BIGINT,
    total_amount BIGINT,
    note         VARCHAR
)
WITH (format = 'ORC')
"""

def load_daily_sales(conn, sh):
    print("\n[일별매출] 시작")
    cur = conn.cursor()

    execute(cur, DDL_DAILY, "테이블 생성(없을 경우)")

    existing = get_existing_dates(cur, TABLE_DAILY)
    print(f"  기존 날짜: {len(existing)}건")

    ws  = sh.worksheet("일별매출")
    rows = ws.get_all_records()

    new_rows = []
    for row in rows:
        date = str(row.get("날짜", "") or "").strip()
        if not date or date in existing:
            continue
        off  = to_int(row.get("OFF거래액", 0))
        on   = to_int(row.get("ON거래액",  0))
        note = str(row.get("특이사항", "") or "").replace("'", "''")
        new_rows.append(f"('{date}', {off}, {on}, {off+on}, '{note}')")

    if not new_rows:
        print("  신규 데이터 없음")
        return

    # 500행씩 배치 INSERT
    batch_size = 500
    inserted = 0
    for i in range(0, len(new_rows), batch_size):
        batch = new_rows[i:i+batch_size]
        sql = f"INSERT INTO {TABLE_DAILY} VALUES {', '.join(batch)}"
        cur.execute(sql)
        inserted += len(batch)

    print(f"  적재 완료: {inserted}행")


# ── 2. 상품별매출 ─────────────────────────────────────

DDL_PRODUCT = f"""
CREATE TABLE IF NOT EXISTS {TABLE_PRODUCT} (
    sale_date    VARCHAR,
    product_code VARCHAR,
    product_name VARCHAR,
    color        VARCHAR,
    size         VARCHAR,
    barcode      VARCHAR,
    unit_price   BIGINT,
    qty          BIGINT,
    amount       BIGINT
)
WITH (format = 'ORC')
"""

def load_product_sales(conn, sh):
    print("\n[상품별매출] 시작")
    cur = conn.cursor()

    execute(cur, DDL_PRODUCT, "테이블 생성(없을 경우)")

    existing = get_existing_dates(cur, TABLE_PRODUCT)
    print(f"  기존 날짜: {len(existing)}건")

    ws   = sh.worksheet("상품별매출")
    rows = ws.get_all_records()

    new_rows = []
    for row in rows:
        date = str(row.get("판매일자", "") or "").strip()
        if not date or date in existing:
            continue
        code    = str(row.get("상품코드", "") or "").replace("'", "''")
        name    = str(row.get("상품명",   "") or "").replace("'", "''")
        color   = str(row.get("칼라명",   "") or "").replace("'", "''")
        size    = str(row.get("사이즈명", "") or "").replace("'", "''")
        barcode = str(row.get("자사바코드","") or "").replace("'", "''")
        price   = to_int(row.get("판매단가",   0))
        qty     = to_int(row.get("판매수량",   0))
        amount  = to_int(row.get("실판매금액", 0))
        new_rows.append(f"('{date}', '{code}', '{name}', '{color}', '{size}', '{barcode}', {price}, {qty}, {amount})")

    if not new_rows:
        print("  신규 데이터 없음")
        return

    batch_size = 500
    inserted = 0
    for i in range(0, len(new_rows), batch_size):
        batch = new_rows[i:i+batch_size]
        sql = f"INSERT INTO {TABLE_PRODUCT} VALUES {', '.join(batch)}"
        cur.execute(sql)
        inserted += len(batch)

    print(f"  적재 완료: {inserted}행")


# ── 메인 ─────────────────────────────────────────────

def main():
    print(f"=== Presto 적재 시작 {datetime.now().strftime('%Y-%m-%d %H:%M')} ===")

    try:
        conn = get_presto_conn()
        print(f"Presto 연결 성공: {PRESTO_HOST}:{PRESTO_PORT}")
    except Exception as e:
        print(f"Presto 연결 실패: {e}")
        sys.exit(1)

    try:
        gs = get_gspread_client()
        sh = gs.open_by_key(SPREADSHEET_ID)
        print("Google Sheets 연결 성공")
    except Exception as e:
        print(f"Google Sheets 연결 실패: {e}")
        sys.exit(1)

    load_daily_sales(conn, sh)
    load_product_sales(conn, sh)

    conn.close()
    print(f"\n=== 완료 {datetime.now().strftime('%Y-%m-%d %H:%M')} ===")


if __name__ == "__main__":
    main()
