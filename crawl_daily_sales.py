"""
플레이엠디 - 영업관리 > 현황 > 매장 일별판매집계표
→ Google Sheets 일별매출 탭에 자동 적재
"""
import asyncio
import json
import os
import getpass
from datetime import datetime, timedelta

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
SPREADSHEET_ID = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"
SHEET_NAME     = "일별매출"
# GitHub Actions: 환경변수 GOOGLE_CREDENTIALS (JSON 문자열)
# 로컬: google_credentials.json 파일
GOOGLE_CREDS_ENV  = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE = "google_credentials.json"

# ── 조회 날짜 (기본: 어제) ────────────────────────────
yesterday  = datetime.today() - timedelta(days=1)
START_DATE = yesterday.strftime("%Y-%m-%d")
END_DATE   = yesterday.strftime("%Y-%m-%d")

LOGIN_URL = "https://playmd.xmd.co.kr/"
# ──────────────────────────────────────────────────────


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


def append_to_sheet(headers, rows):
    client = get_gspread_client()
    sh     = client.open_by_key(SPREADSHEET_ID)
    ws     = sh.worksheet(SHEET_NAME)

    existing = ws.get_all_values()

    # 헤더가 없으면 첫 행에 추가
    if not existing:
        ws.append_row(["수집일자"] + headers)

    for row in rows:
        ws.append_row([START_DATE] + row)

    print(f"Google Sheets 적재 완료: {len(rows)}행")


async def run():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page    = await browser.new_page()

        # 1) 로그인
        print("로그인 중...")
        await page.goto(LOGIN_URL)
        await page.wait_for_load_state("networkidle")

        # 텍스트 입력창: 첫 번째=회원사ID, 두 번째=하위ID
        text_inputs = page.locator("input[type='text'], input[type='id']")
        await text_inputs.nth(0).fill(COMPANY_USER)
        await text_inputs.nth(1).fill(USERNAME)

        # 비밀번호 입력창: 첫 번째=회원사PW, 두 번째=하위PW
        pw_inputs = page.locator("input[type='password']")
        await pw_inputs.nth(0).fill(COMPANY_PASS)
        await pw_inputs.nth(1).fill(PASSWORD)

        await page.click("button[type='submit'], input[type='submit'], .btn-login, button:has-text('로그인')")
        await page.wait_for_load_state("networkidle")
        print(f"로그인 후 URL: {page.url}")

        # 2) 메뉴 탐색
        print("메뉴 이동 중...")
        await page.click("text=영업관리")
        await page.wait_for_timeout(500)
        await page.click("text=현황")
        await page.wait_for_timeout(500)
        await page.click("text=매장 일별판매집계표")
        await page.wait_for_load_state("networkidle")
        print(f"페이지: {page.url}")

        # 3) 날짜 필터
        print(f"기간: {START_DATE} ~ {END_DATE}")
        try:
            start_inputs = page.locator("input[type='date'], input[placeholder*='시작'], input[id*='start'], input[name*='start']")
            end_inputs   = page.locator("input[type='date'], input[placeholder*='종료'], input[id*='end'], input[name*='end']")

            if await start_inputs.count() > 0:
                await start_inputs.first.fill(START_DATE)
            if await end_inputs.count() > 1:
                await end_inputs.nth(1).fill(END_DATE)
            elif await end_inputs.count() > 0:
                await end_inputs.first.fill(END_DATE)

            await page.click("button:has-text('조회'), button:has-text('검색'), input[value='조회']")
            await page.wait_for_load_state("networkidle")
        except Exception as e:
            print(f"날짜 필터 오류: {e}")
            await page.screenshot(path="debug_filter.png", full_page=True)

        await page.screenshot(path="debug_table.png", full_page=True)

        # 4) 테이블 추출
        print("데이터 추출 중...")
        headers = await page.locator("table thead th, table thead td").all_text_contents()
        if not headers:
            headers = await page.locator("table tr:first-child th").all_text_contents()
        headers = [h.strip() for h in headers if h.strip()]
        print(f"헤더: {headers}")

        tr_elements = page.locator("table tbody tr")
        count = await tr_elements.count()
        print(f"행 수: {count}")

        rows = []
        for i in range(count):
            cells = await tr_elements.nth(i).locator("td").all_text_contents()
            cells = [c.strip() for c in cells]
            if any(cells):
                rows.append(cells)

        await browser.close()

        # 5) Google Sheets 적재
        if rows:
            append_to_sheet(headers, rows)
        else:
            print("데이터 없음. debug_table.png 확인 필요.")


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
