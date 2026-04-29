"""
플레이엠디 - 매장 일별판매집계표
1. Playwright로 로그인 → 쿠키 추출
2. requests로 API 직접 호출
3. 어제 일별 거래액 → Google Sheets 적재
"""
import asyncio
import json
import os
import getpass
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
USERNAME      = os.environ.get("PLAYMD_USER", "")
PASSWORD      = os.environ.get("PLAYMD_PASS", "")

# ── Google Sheets 설정 ────────────────────────────────
SPREADSHEET_ID   = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"
SHEET_NAME       = "일별매출"
GOOGLE_CREDS_ENV = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE = "google_credentials.json"

# ── 조회 기준: 어제 ───────────────────────────────────
yesterday = datetime.today() - timedelta(days=1)
YYMM      = yesterday.strftime("%Y%m")   # 예: "202604"
DAY_KEY   = f"D{yesterday.day}"          # 예: "D28"

LOGIN_URL  = "https://playmd.xmd.co.kr/"
API_URL    = "https://playmd.xmd.co.kr/api/xsal/xsal6020q_s03_peace"


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


def upload_to_sheets(rows: list):
    client = get_gspread_client()
    sh = client.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet(SHEET_NAME)

    existing = ws.get_all_values()
    if not existing:
        ws.append_row(["날짜", "매장코드", "매장명", "일별거래액"])

    for row in rows:
        ws.append_row(row)

    print(f"Google Sheets 적재 완료: {len(rows)}행")


async def login_and_get_cookies() -> dict:
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

        # 쿠키 추출
        cookies = await page.context.cookies()
        await browser.close()

        cookie_dict = {c["name"]: c["value"] for c in cookies}
        print(f"추출된 쿠키: {list(cookie_dict.keys())}")
        return cookie_dict


def fetch_sales_data(cookies: dict) -> list:
    payload = {
        "AGT_DESC":     {},
        "I_AGTSUB":     "",
        "I_CAGTAREA":   "",
        "I_CAGTCD":     "",
        "I_CAGTCST":    "",
        "I_COUPONYN":   "Y",
        "I_DISCOUNTYN": "Y",
        "I_GROUP":      "1",
        "I_MILEYN":     "Y",
        "I_PRIME":      "",
        "I_SELNM":      "",
        "I_TAG":        "1",
        "I_USEYN":      "Y",
        "I_VZ21":       "",
        "I_YYMM":       YYMM,
        "WHERE_DESC":   {}
    }

    headers = {
        "Content-Type": "application/json",
        "Referer": "https://playmd.xmd.co.kr/xsal/xsal6020q/xsal6020q.html",
        "Origin":  "https://playmd.xmd.co.kr",
    }

    resp = requests.post(API_URL, json=payload, headers=headers, cookies=cookies, timeout=30)
    print(f"API 응답: {resp.status_code}")
    data = resp.json()
    print(f"데이터 행 수: {len(data)}")

    # 어제 날짜 컬럼(DAY_KEY) 추출
    date_str = yesterday.strftime("%Y-%m-%d")
    rows = []
    for item in data:
        store_cd   = item.get("AGTCD", "")
        store_nm   = item.get("AGTNM", "")
        day_amount = item.get(DAY_KEY, "")
        if store_cd:
            rows.append([date_str, store_cd, store_nm, day_amount])

    print(f"어제({DAY_KEY}) 데이터: {len(rows)}행")
    return rows


async def main():
    cookies = await login_and_get_cookies()
    rows    = fetch_sales_data(cookies)
    if rows:
        upload_to_sheets(rows)
    else:
        print("데이터 없음")


if __name__ == "__main__":
    if not COMPANY_USER:
        COMPANY_USER = input("회원사 ID: ")
    if not COMPANY_PASS:
        COMPANY_PASS = getpass.getpass("회원사 PW: ")
    if not USERNAME:
        USERNAME = input("하위 ID: ")
    if not PASSWORD:
        PASSWORD = getpass.getpass("하위 PW: ")
    asyncio.run(main())
