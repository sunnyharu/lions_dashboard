"""
플레이엠디 매장판매일보 → 상품별 집계 → Google Sheets
1. Playwright로 로그인 → 매장판매일보 이동
2. 어제 날짜 조회 → 엑셀 다운로드
3. pandas로 상품별 집계
4. Google Sheets '상품별매출' 시트 적재
"""
import asyncio
import io
import json
import os
from datetime import datetime, timedelta

import pandas as pd
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
DATE_PARAM = yesterday.strftime("%Y-%m-%d")   # 2026-05-20
DATE_SHEET = yesterday.strftime("%Y.%m.%d")   # 2026.05.20

LOGIN_URL = "https://playmd.xmd.co.kr/"


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
    """집계된 상품별 데이터 Google Sheets 적재 (중복 날짜 스킵)"""
    if not rows:
        print("업로드할 데이터 없음")
        return

    client = get_gspread_client()
    sh     = client.open_by_key(SPREADSHEET_ID)

    try:
        ws = sh.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=5000, cols=10)
        ws.append_row(["날짜", "상품코드", "상품명", "칼라명", "사이즈명", "판매수량", "판매금액"])
        print(f"시트 '{SHEET_NAME}' 생성")

    existing = ws.get_all_values()
    header   = existing[0] if existing else []

    # 이미 같은 날짜 데이터가 있으면 스킵
    existing_dates = {r[0].strip() for r in existing[1:] if r}
    if DATE_SHEET in existing_dates:
        print(f"이미 존재하는 날짜 스킵: {DATE_SHEET}")
        return

    ws.append_rows(rows, value_input_option="USER_ENTERED")
    print(f"Google Sheets 적재 완료: {len(rows)}행 삽입")


# ── 엑셀 파싱 & 집계 ─────────────────────────────────

def parse_and_aggregate(excel_bytes: bytes) -> list:
    """엑셀 파일 읽어서 상품별로 집계"""
    df = pd.read_excel(io.BytesIO(excel_bytes), engine="openpyxl")
    print(f"엑셀 로드: {len(df)}행, 컬럼: {list(df.columns)}")

    # 컬럼명 공백 제거
    df.columns = df.columns.str.strip()

    # 상품코드·상품명 컬럼 찾기 (유사 이름 대응)
    col_map = {}
    for col in df.columns:
        c = col.strip()
        if "상품코드" in c:   col_map["상품코드"] = col
        elif "상품명" in c:   col_map["상품명"]   = col
        elif "칼라명" in c:   col_map["칼라명"]   = col
        elif "사이즈명" in c: col_map["사이즈명"] = col
        elif c in ("수량", "판매수량", "QTY"):   col_map["수량"]   = col
        elif c in ("금액", "판매금액", "매출금액", "거래금액"): col_map["금액"] = col

    print(f"컬럼 매핑: {col_map}")

    required = ["상품코드", "상품명"]
    for r in required:
        if r not in col_map:
            print(f"필수 컬럼 '{r}' 없음 → 집계 불가")
            return []

    # 집계 기준 컬럼
    group_cols = [col_map[k] for k in ["상품코드", "상품명", "칼라명", "사이즈명"] if k in col_map]

    # 숫자 변환
    for key in ["수량", "금액"]:
        if key in col_map:
            df[col_map[key]] = (
                df[col_map[key]]
                .astype(str)
                .str.replace(",", "")
                .str.strip()
            )
            df[col_map[key]] = pd.to_numeric(df[col_map[key]], errors="coerce").fillna(0)

    agg_dict = {}
    if "수량" in col_map: agg_dict[col_map["수량"]] = "sum"
    if "금액" in col_map: agg_dict[col_map["금액"]] = "sum"

    if not agg_dict:
        print("수량/금액 컬럼 없음 → 집계 불가")
        return []

    agg = df.groupby(group_cols, as_index=False).agg(agg_dict)
    agg = agg.sort_values(col_map.get("금액", group_cols[0]), ascending=False)

    print(f"집계 결과: {len(agg)}개 상품")

    rows = []
    for _, r in agg.iterrows():
        row = [DATE_SHEET]
        row.append(str(r.get(col_map.get("상품코드", ""), "")).strip())
        row.append(str(r.get(col_map.get("상품명",   ""), "")).strip())
        row.append(str(r.get(col_map.get("칼라명",   ""), "")).strip()   if "칼라명"   in col_map else "")
        row.append(str(r.get(col_map.get("사이즈명", ""), "")).strip()   if "사이즈명" in col_map else "")
        row.append(int(r.get(col_map.get("수량", ""),  0)) if "수량" in col_map else 0)
        row.append(int(r.get(col_map.get("금액", ""),  0)) if "금액" in col_map else 0)
        rows.append(row)

    return rows


# ── Playwright 자동화 ────────────────────────────────

async def crawl() -> bytes | None:
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(accept_downloads=True)
        page    = await context.new_page()

        # ── 로그인 ──────────────────────────────────
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

        # ── 매장판매일보 메뉴 검색 ───────────────────
        print("매장판매일보 메뉴 이동 중...")
        await page.screenshot(path="debug_01_login.png")

        search_input = page.locator("input[placeholder*='메뉴검색'], input[placeholder*='메뉴'], .menu-search input").first
        await search_input.fill("매장판매일보")
        await page.wait_for_timeout(1000)

        # 검색 결과에서 클릭
        menu_item = page.locator("text=매장판매일보").first
        await menu_item.wait_for(timeout=5000)
        await menu_item.click()
        await page.wait_for_timeout(2000)
        await page.wait_for_load_state("networkidle")
        await page.screenshot(path="debug_02_menu.png")
        print(f"메뉴 이동 완료: {page.url}")

        # ── 날짜 설정 (어제) ─────────────────────────
        print(f"날짜 설정: {DATE_PARAM}")

        # 판매기간 시작일 / 종료일 모두 어제로 설정
        date_inputs = page.locator("input[type='text']").all()
        date_inputs = await page.locator("input[type='text']").all()

        # 날짜 입력 필드 찾기 (값이 날짜 형식인 것)
        for inp in date_inputs[:5]:
            val = await inp.input_value()
            if "-" in val and len(val) == 10:  # YYYY-MM-DD 형식
                await inp.triple_click()
                await inp.fill(DATE_PARAM)
                await page.wait_for_timeout(300)

        await page.screenshot(path="debug_03_date.png")

        # ── 조회 클릭 ────────────────────────────────
        print("조회 클릭...")
        await page.locator("button:has-text('조회'), a:has-text('조회')").first.click()
        await page.wait_for_timeout(5000)
        await page.wait_for_load_state("networkidle")
        await page.screenshot(path="debug_04_result.png")
        print("조회 완료")

        # ── 엑셀 다운로드 ────────────────────────────
        print("엑셀 다운로드 중...")
        async with page.expect_download(timeout=30000) as dl_info:
            await page.locator("button:has-text('엑셀'), a:has-text('엑셀')").first.click()

        download    = await dl_info.value
        excel_bytes = await download.read()
        print(f"엑셀 다운로드 완료: {len(excel_bytes):,} bytes")

        await browser.close()
        return excel_bytes


# ── 메인 ────────────────────────────────────────────

async def main():
    excel_bytes = await crawl()
    if not excel_bytes:
        print("엑셀 다운로드 실패")
        return

    rows = parse_and_aggregate(excel_bytes)
    if rows:
        upload_to_sheets(rows)
    else:
        print("집계 데이터 없음")


if __name__ == "__main__":
    asyncio.run(main())
