"""
PlayMD 상품별 매출 백필 스크립트
기간 전체를 단일 API 호출로 가져와 날짜+상품 기준으로 그룹핑 후 Google Sheets 적재
"""
import asyncio
import json
import os
import requests
from datetime import datetime, timedelta

from dotenv import load_dotenv
from playwright.async_api import async_playwright
import gspread
from google.oauth2.service_account import Credentials

load_dotenv()

# ── 백필 기간 설정 ─────────────────────────────────────
FROM_DATE = "20260316"   # 시작일
TO_DATE   = "20260521"   # 종료일 (어제)

FROM_DISPLAY = "2026.03.16"
TO_DISPLAY   = "2026.05.21"

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

LOGIN_URL = "https://playmd.xmd.co.kr/"
API_URL   = "https://playmd.xmd.co.kr/api/xagt/xagt5000q_ver2_s01"

SHEET_HEADER = ["판매일자", "상품코드", "상품명", "칼라명", "사이즈명", "자사바코드", "판매단가", "판매수량", "실판매금액"]

KEY_MAP = {
    "판매일자":  ["ASALDT", "SALDT",  "SALEDT"],
    "상품코드":  ["GODCD",  "ITEMCD", "PRODCD"],
    "상품명":    ["GODNM",  "ITEMNM", "PRODNM"],
    "칼라명":    ["CRNM",   "COLORNM","CLRNM"],
    "사이즈명":  ["SZNM",   "SIZENM", "SIZNM"],
    "자사바코드":["BARNO1", "BARCD",  "BARCODE", "MAINBARCD"],
    "판매단가":  ["SALPR",  "SCHPR",  "PRICE",   "GODPR",  "SALUPRC"],
    "판매수량":  ["SALQT",  "QTY",    "SALQTY"],
    "실판매금액":["RSALAMT","SALAMT", "NETAMT",  "RSLAMT", "ACSLAMT"],
}


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


def fetch_all(cookies: dict) -> list:
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
        "I_FROMDATE": FROM_DATE,
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
        "I_TODATE":  TO_DATE,
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

    print(f"API 호출: {FROM_DATE} ~ {TO_DATE}")
    resp = requests.post(API_URL, json=payload, headers=headers, cookies=cookies, timeout=120)
    print(f"API 응답: {resp.status_code}, {len(resp.text)}자")
    if resp.status_code != 200:
        print(f"오류: {resp.text[:300]}")
        return []

    data = resp.json()
    print(f"전체 행 수: {len(data)}행")
    return data if isinstance(data, list) else []


def aggregate_by_date(data: list) -> list:
    if not data:
        return []

    sample = data[0] if isinstance(data[0], dict) else {}

    def find_key(candidates):
        for c in candidates:
            if c in sample:
                return c
        return None

    col_keys = {col: find_key(cands) for col, cands in KEY_MAP.items()}
    print("컬럼 매핑:")
    for col, k in col_keys.items():
        print(f"  {col} → {k or '없음'}")

    def to_int(v):
        try: return int(float(str(v).replace(",", "").strip()))
        except: return 0

    def fmt_date(v):
        # ASALDT: "20260316" → "2026.03.16"
        s = str(v).strip()
        if len(s) == 8 and s.isdigit():
            return f"{s[:4]}.{s[4:6]}.{s[6:]}"
        return s

    # 날짜 + 상품SKU 기준 집계
    agg = {}
    for row in data:
        if not isinstance(row, dict):
            continue

        raw_date = row.get(col_keys.get("판매일자") or "", "")
        date_sheet = fmt_date(raw_date)

        key = (
            date_sheet,
            str(row.get(col_keys.get("상품코드")  or "", "") or "").strip(),
            str(row.get(col_keys.get("상품명")    or "", "") or "").strip(),
            str(row.get(col_keys.get("칼라명")    or "", "") or "").strip(),
            str(row.get(col_keys.get("사이즈명")  or "", "") or "").strip(),
            str(row.get(col_keys.get("자사바코드") or "", "") or "").strip(),
        )
        단가_raw = row.get(col_keys.get("판매단가")  or "", 0) or 0
        수량_raw = row.get(col_keys.get("판매수량")  or "", 0) or 0
        금액_raw = row.get(col_keys.get("실판매금액") or "", 0) or 0

        if key not in agg:
            agg[key] = {"단가": to_int(단가_raw), "수량": 0, "금액": 0}
        agg[key]["수량"] += to_int(수량_raw)
        agg[key]["금액"] += to_int(금액_raw)

    # 날짜 오름차순, 같은 날은 금액 내림차순
    rows = []
    for (날짜, 코드, 명, 칼라, 사이즈, 바코드), v in sorted(agg.items(), key=lambda x: (x[0][0], -x[0][1]["금액"]) if False else (x[0][0], -x[1]["금액"])):
        rows.append([날짜, 코드, 명, 칼라, 사이즈, 바코드, v["단가"], v["수량"], v["금액"]])

    dates = sorted({r[0] for r in rows})
    print(f"집계 결과: {len(rows)}행 ({len(dates)}일치, {dates[0]} ~ {dates[-1]})")
    return rows


def upload_to_sheets(rows: list):
    if not rows:
        print("업로드할 데이터 없음")
        return

    client = get_gspread_client()
    sh     = client.open_by_key(SPREADSHEET_ID)

    try:
        ws = sh.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=100000, cols=len(SHEET_HEADER))

    existing = ws.get_all_values()
    if not existing or existing[0] != SHEET_HEADER:
        ws.insert_row(SHEET_HEADER, 1)
        print("헤더 복구 완료")
        existing = ws.get_all_values()

    existing_dates = {r[0].strip() for r in existing[1:] if r}
    print(f"시트 기존 날짜: {sorted(existing_dates)}")

    new_rows = [r for r in rows if r[0] not in existing_dates]
    skip_dates = {r[0] for r in rows if r[0] in existing_dates}

    if skip_dates:
        print(f"스킵 날짜 ({len(skip_dates)}일): {sorted(skip_dates)}")
    if not new_rows:
        print("모든 날짜가 이미 존재하여 업로드 없음")
        return

    ws.append_rows(new_rows, value_input_option="USER_ENTERED")
    new_dates = sorted({r[0] for r in new_rows})
    print(f"적재 완료: {len(new_rows)}행, {len(new_dates)}일치 ({new_dates[0]} ~ {new_dates[-1]})")


async def main():
    cookies = await get_cookies()
    data    = fetch_all(cookies)
    if not data:
        print("데이터 없음")
        return
    rows = aggregate_by_date(data)
    upload_to_sheets(rows)


if __name__ == "__main__":
    asyncio.run(main())
