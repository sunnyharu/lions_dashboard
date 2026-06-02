"""
플레이엠디 매장판매일보 → 상품별 집계 → Google Sheets
1. Playwright로 로그인 → 쿠키 추출
2. requests로 API 직접 호출 (xagt5000q_ver2_s01)
3. JSON 파싱 → 상품별 집계
4. Google Sheets '상품별매출' 시트 적재
"""
import asyncio
import json
import os
import time
import requests
from datetime import datetime, timedelta

from dotenv import load_dotenv
from playwright.async_api import async_playwright
import gspread
from google.oauth2.service_account import Credentials

load_dotenv()

# ── 플레이엠디 계정 ────────────────────────────────────
COMPANY_USER = os.environ.get("PLAYMD_COMPANY_USER", "")
COMPANY_PASS = os.environ.get("PLAYMD_COMPANY_PASS", "")
USERNAME     = os.environ.get("PLAYMD_USER", "")
PASSWORD     = os.environ.get("PLAYMD_PASS", "")

# ── Google Sheets 설정 ────────────────────────────────
SPREADSHEET_ID    = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"
SHEET_NAME        = "상품별매출"
GOOGLE_CREDS_ENV  = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE = "google_credentials.json"

# ── 조회 기준: 어제 ───────────────────────────────────
yesterday  = datetime.today() - timedelta(days=1)
DATE_PARAM = yesterday.strftime("%Y%m%d")   # 20260520
DATE_SHEET = yesterday.strftime("%Y.%m.%d") # 2026.05.20

LOGIN_URL = "https://playmd.xmd.co.kr/"
API_URL   = "https://playmd.xmd.co.kr/api/xagt/xagt5000q_ver2_s01"

SHEET_HEADER = ["판매일자", "상품코드", "상품명", "칼라명", "사이즈명", "자사바코드", "판매단가", "판매수량", "실판매금액"]


# ── Google Sheets ────────────────────────────────────

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


def upload_to_sheets(rows: list):
    if not rows:
        print("업로드할 데이터 없음")
        return

    # Google Sheets API 503 등 일시 오류 시 최대 3회 재시도
    for attempt in range(3):
        try:
            client = get_gspread_client()
            sh     = client.open_by_key(SPREADSHEET_ID)
            break
        except Exception as e:
            if attempt < 2:
                print(f"Sheets 연결 오류 (재시도 {attempt+1}/3): {e}")
                time.sleep(10)
            else:
                raise

    try:
        ws = sh.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=50000, cols=len(SHEET_HEADER))
        ws.append_row(SHEET_HEADER)
        print(f"시트 '{SHEET_NAME}' 생성")

    existing = ws.get_all_values()

    # 헤더 없거나 잘못됐으면 복구
    if not existing or existing[0] != SHEET_HEADER:
        ws.insert_row(SHEET_HEADER, 1)
        print("헤더 복구 완료")
        existing = ws.get_all_values()

    existing_dates = {r[0].strip() for r in existing[1:] if r}
    if DATE_SHEET in existing_dates:
        print(f"이미 존재하는 날짜 스킵: {DATE_SHEET}")
        return

    ws.append_rows(rows, value_input_option="USER_ENTERED")
    print(f"Google Sheets 적재 완료: {len(rows)}행 삽입")


# ── Playwright 로그인 (쿠키만 추출) ──────────────────

async def get_cookies() -> dict:
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page    = await browser.new_page()

        print("로그인 중...")
        await page.goto(LOGIN_URL)
        await page.wait_for_load_state("networkidle")
        await page.fill("#txt-tenantLoginId", COMPANY_USER)
        await page.fill("#pw-tenantPassword", COMPANY_PASS)
        await page.fill("#txt-userLoginId",   USERNAME)
        await page.fill("#txt-userPassword",  PASSWORD)
        await page.locator("button:has-text('플레이엠디 로그인')").click()
        await page.wait_for_timeout(3000)
        await page.wait_for_load_state("networkidle")
        print(f"로그인 완료: {page.url}")

        cookies = await page.context.cookies()
        await browser.close()

        return {c["name"]: c["value"] for c in cookies}


# ── API 호출 ─────────────────────────────────────────

def fetch_product_data(cookies: dict) -> list:
    payload = {
        "AGT_DESC":              {},
        "WHERE_DESC":            {},
        "conditionList":         {"viewType": "H"},
        "hideOnlinePrivacyInfo": True,
        "mergeCell":             False,
        "productCodeBlock":      {"PMSTCD1": False, "PMSTCD2": False, "PMSTCD3": False},
        "viewAgtSubsum":         False,
        "viewDtSubsum":          False,
        "viewImage":             False,
        "viewMainBarcode":       True,
        "viewOnlineInfo":        False,
        "viewOnlineInfoDisabled": True,
        "viewReceiptSubsum":     False,
        "viewSetInfo":           False,
        "viewSetInfoDisabled":   False,
        "viewSetSalse":          False,
        "viewSubBarcode1":       False,
        "viewSubBarcode2":       False,
        "viewTAXFREEInfo":       False,
        "ACSTNAME":  "",
        "CSTCD2NM":  "",
        "CSTCDNM":   "",
        "GODNM":     "",
        "I_ACSTNO":  "",
        "I_AGTCD":   "",
        "I_CSTCD":   "",
        "I_CSTCD2":  "",
        "I_DESIGNER": "",
        "I_EVTGB":   "",
        "I_FROMDATE": DATE_PARAM,
        "I_GODCD":   "",
        "I_INID":    "",
        "I_MEMCD":   "kakaoent",
        "I_NOTE":    "",
        "I_PLANGB":  "",
        "I_REVENUE": "",
        "I_SALESMAN": "",
        "I_SALGB":   "1",
        "I_SALNM":   "",
        "I_SGB":     [],
        "I_TEAM":    "",
        "I_TODATE":  DATE_PARAM,
        "I_USERCST": "03",
        "I_USRGUBN": "1",
    }

    headers = {
        "Accept":        "application/json, text/plain, */*",
        "Content-Type":  "application/json;charset=UTF-8",
        "Origin":        "https://playmd.xmd.co.kr",
        "Referer":       "https://playmd.xmd.co.kr/xagt/xagt5000q/xagt5000q.html",
        "X-Current-Url": "https://playmd.xmd.co.kr/xagt/xagt5000q/xagt5000q.html",
        "Xmd-Session":   cookies.get("xmd_session", ""),
        "User-Agent":    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36",
    }

    resp = requests.post(API_URL, json=payload, headers=headers, cookies=cookies, timeout=60)
    print(f"API 응답: {resp.status_code}")
    if resp.status_code != 200:
        print(f"오류: {resp.text[:300]}")
        return []

    data = resp.json()
    print(f"응답 데이터 수: {len(data)}행")
    if data:
        print(f"첫 번째 행 키: {list(data[0].keys()) if isinstance(data[0], dict) else data[0]}")

    return data


# ── 집계 ────────────────────────────────────────────

def aggregate(data: list) -> list:
    if not data:
        return []

    # 첫 번째 행으로 컬럼 파악
    sample = data[0] if isinstance(data[0], dict) else {}
    print(f"컬럼 목록: {list(sample.keys())}")
    print(f"첫 번째 행 전체 데이터:")
    for k, v in sample.items():
        print(f"  {k}: {v}")

    # 컬럼명 매핑 (PlayMD 내부 키 → 우리 컬럼명)
    KEY_MAP = {
        "판매일자":  ["ASALDT", "SALDT",  "SALEDT", "SALDATE"],
        "상품코드":  ["GODCD",  "ITEMCD", "PRODCD"],
        "상품명":    ["GODNM",  "ITEMNM", "PRODNM"],
        "칼라명":    ["CRNM",   "COLORNM","CLRNM"],
        "사이즈명":  ["SZNM",   "SIZENM", "SIZNM"],
        "자사바코드":["BARNO1", "BARCD",  "BARCODE", "MAINBARCD"],
        "판매단가":  ["SALPR",  "SCHPR",  "PRICE",  "GODPR",  "SALUPRC", "SAPRC", "SLPRC"],
        "판매수량":  ["SALQT",  "QTY",    "SALQTY", "SALQTA"],
        "실판매금액":["RSALAMT","SALAMT", "NETAMT", "RSLAMT", "ACSLAMT", "NETSLAMT", "SLAMTI", "ACSALAMT"],
    }

    def find_key(row, candidates):
        for c in candidates:
            if c in row:
                return c
        return None

    # 컬럼 키 확정
    col_keys = {}
    for col, candidates in KEY_MAP.items():
        k = find_key(sample, candidates)
        if k:
            col_keys[col] = k
            print(f"  {col} → {k}")
        else:
            print(f"  {col} → 키 없음 (후보: {candidates})")

    # 집계 (상품코드+상품명+칼라명+사이즈명+자사바코드+판매단가 기준)
    agg = {}
    for row in data:
        if not isinstance(row, dict):
            continue

        key = (
            str(row.get(col_keys.get("상품코드",  ""), "") or "").strip(),
            str(row.get(col_keys.get("상품명",    ""), "") or "").strip(),
            str(row.get(col_keys.get("칼라명",    ""), "") or "").strip(),
            str(row.get(col_keys.get("사이즈명",  ""), "") or "").strip(),
            str(row.get(col_keys.get("자사바코드",""), "") or "").strip(),
        )
        단가_raw = row.get(col_keys.get("판매단가", ""), 0) or 0
        수량_raw = row.get(col_keys.get("판매수량", ""), 0) or 0
        금액_raw = row.get(col_keys.get("실판매금액", ""), 0) or 0

        def to_int(v):
            try: return int(float(str(v).replace(",", "").strip()))
            except: return 0

        if key not in agg:
            agg[key] = {"단가": to_int(단가_raw), "수량": 0, "금액": 0}
        agg[key]["수량"] += to_int(수량_raw)
        agg[key]["금액"] += to_int(금액_raw)

    rows = []
    for (코드, 명, 칼라, 사이즈, 바코드), v in sorted(agg.items(), key=lambda x: -x[1]["금액"]):
        rows.append([DATE_SHEET, 코드, 명, 칼라, 사이즈, 바코드, v["단가"], v["수량"], v["금액"]])

    print(f"집계 결과: {len(rows)}개 상품 SKU")
    return rows


# ── 메인 ────────────────────────────────────────────

async def main():
    cookies = await get_cookies()
    data    = fetch_product_data(cookies)
    if not data:
        print("데이터 없음")
        return
    rows = aggregate(data)
    if rows:
        upload_to_sheets(rows)


if __name__ == "__main__":
    asyncio.run(main())
