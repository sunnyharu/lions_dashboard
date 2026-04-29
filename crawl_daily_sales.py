"""
플레이엠디 - 영업관리 > 현황 > 매장 일별판매집계표
엑셀 다운로드 → 일별 거래액 파싱 → Google Sheets 적재
"""
import asyncio
import json
import os
import getpass
import tempfile
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
USERNAME      = os.environ.get("PLAYMD_USER", "")
PASSWORD      = os.environ.get("PLAYMD_PASS", "")

# ── Google Sheets 설정 ────────────────────────────────
SPREADSHEET_ID    = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"
SHEET_NAME        = "일별매출"
GOOGLE_CREDS_ENV  = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE = "google_credentials.json"

# ── 조회 기준: 어제 ───────────────────────────────────
yesterday  = datetime.today() - timedelta(days=1)
YEAR       = str(yesterday.year)
MONTH      = f"{yesterday.month:02d}"
DAY        = f"{yesterday.day:02d}"

LOGIN_URL = "https://playmd.xmd.co.kr/"


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


def parse_excel_and_upload(filepath: str):
    df = pd.read_excel(filepath, header=None)
    print(f"엑셀 로드 완료: {df.shape}")
    print(df.head(10).to_string())

    # 헤더 행 찾기 (매장명 또는 날짜 숫자가 있는 행)
    header_row = None
    for i, row in df.iterrows():
        row_str = " ".join(str(v) for v in row.values)
        if "매장명" in row_str or "매장코드" in row_str:
            header_row = i
            break

    if header_row is None:
        print("헤더 행을 찾지 못했습니다. 엑셀 구조 확인 필요.")
        return

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    print(f"컬럼: {list(df.columns)}")

    # 어제 날짜 컬럼 찾기 (예: "04", "4", "04(토)" 등)
    day_col = None
    for col in df.columns:
        col_str = str(col).strip()
        if col_str == DAY or col_str == str(int(DAY)) or col_str.startswith(DAY):
            day_col = col
            break

    if day_col is None:
        print(f"날짜 컬럼 '{DAY}' 를 찾지 못했습니다. 컬럼 목록: {list(df.columns)}")
        return

    print(f"어제({DAY}일) 컬럼: '{day_col}'")

    # 소계/합계 행 또는 전체 데이터 추출
    result_rows = []
    date_str = f"{YEAR}-{MONTH}-{DAY}"

    for _, row in df.iterrows():
        store_name = str(row.get("매장명", "")).strip()
        day_value  = row.get(day_col, "")
        if store_name and store_name not in ["nan", ""]:
            result_rows.append([date_str, store_name, day_value])

    print(f"추출된 행: {len(result_rows)}")

    # Google Sheets 적재
    client = get_gspread_client()
    sh     = client.open_by_key(SPREADSHEET_ID)
    ws     = sh.worksheet(SHEET_NAME)

    existing = ws.get_all_values()
    if not existing:
        ws.append_row(["날짜", "매장명", "일별거래액"])

    for row in result_rows:
        ws.append_row(row)

    print(f"Google Sheets 적재 완료: {len(result_rows)}행")


def js_click(text):
    return f"""() => {{
        const links = Array.from(document.querySelectorAll('a'));
        const target = links.find(a => a.textContent.trim() === '{text}' && a.offsetParent !== null);
        if (target) {{ target.click(); return true; }}
        return false;
    }}"""


async def run():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page    = await browser.new_page()

        # 1) 로그인
        print("로그인 중...")
        await page.goto(LOGIN_URL)
        await page.wait_for_load_state("networkidle")

        inputs_info = await page.evaluate("""() => {
            return Array.from(document.querySelectorAll('input')).map(el => ({
                type: el.type, name: el.name, id: el.id,
                placeholder: el.placeholder, visible: el.offsetParent !== null
            }));
        }""")
        print(f"입력 필드 목록: {inputs_info}")

        await page.fill("#txt-tenantLoginId", COMPANY_USER)
        await page.fill("#pw-tenantPassword", COMPANY_PASS)
        await page.fill("#txt-userLoginId",   USERNAME)
        await page.fill("#txt-userPassword",  PASSWORD)

        await page.locator("button:has-text('플레이엠디 로그인')").click()
        await page.wait_for_timeout(3000)
        await page.wait_for_load_state("networkidle")
        print(f"로그인 후 URL: {page.url}")

        # 2) 메뉴 탐색
        print("메뉴 이동 중...")
        await page.wait_for_timeout(2000)

        def js_click_contains(text):
            return f"""() => {{
                const all = Array.from(document.querySelectorAll('a, button, span, li'));
                const target = all.find(el => el.textContent.trim().includes('{text}') && el.offsetParent !== null);
                if (target) {{ target.click(); return true; }}
                return false;
            }}"""

        # 2) 직접 URL 이동 (URL 확인 완료)
        TARGET_URL = "https://playmd.xmd.co.kr/xsal/xsal6020q/xsal6020q.html"
        print(f"페이지 이동: {TARGET_URL}")
        await page.goto(TARGET_URL)
        try:
            await page.wait_for_selector("text=판매년월", timeout=15000)
        except:
            await page.wait_for_timeout(5000)
        await page.screenshot(path="debug_after_menu.png", full_page=True)
        print("매장일별판매집계표 로드 완료")

        # 3) 조회 버튼 클릭 시도 (실패해도 계속 진행)
        print(f"조회 기준: {YEAR}년 {MONTH}월")
        try:
            # 버튼 구조 확인
            btns = await page.evaluate("""() => {
                return Array.from(document.querySelectorAll('button, a, span, div'))
                    .filter(el => el.textContent.trim().includes('조회') && el.offsetParent !== null)
                    .map(el => ({tag: el.tagName, text: el.textContent.trim(), class: el.className}))
                    .slice(0, 5);
            }""")
            print(f"조회 버튼 후보: {btns}")
            await page.locator(":visible").filter(has_text="조회").first.click(timeout=5000)
            await page.wait_for_timeout(3000)
        except Exception as e:
            print(f"조회 클릭 실패(계속 진행): {e}")
        await page.screenshot(path="debug_after_search.png", full_page=True)

        # 4) 엑셀 다운로드 (네이티브 클릭)
        print("엑셀 다운로드 중...")
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name

        async with page.expect_download(timeout=60000) as download_info:
            await page.locator("button:has-text('엑셀'), a:has-text('엑셀')").first.click()
            # 확인 팝업 "예" 버튼 처리
            try:
                await page.wait_for_selector("button:has-text('예'), a:has-text('예')", timeout=5000)
                await page.locator("button:has-text('예'), a:has-text('예')").first.click()
            except:
                pass

        download = await download_info.value
        await download.save_as(tmp_path)
        print(f"다운로드 완료: {tmp_path}")

        await browser.close()

        # 5) 파싱 & 적재
        parse_excel_and_upload(tmp_path)


if __name__ == "__main__":
    if not COMPANY_USER:
        COMPANY_USER = input("회원사 ID: ")
    if not COMPANY_PASS:
        COMPANY_PASS = getpass.getpass("회원사 PW: ")
    if not USERNAME:
        USERNAME = input("하위 ID: ")
    if not PASSWORD:
        PASSWORD = getpass.getpass("하위 PW: ")
    asyncio.run(run())
