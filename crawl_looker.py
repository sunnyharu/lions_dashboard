"""
Google Looker Studio 대시보드에서 온라인 매출 데이터 스크래핑
→ Google Sheets '온라인매출' 탭에 적재
"""
import asyncio
import json
import os
import re
import time
from datetime import datetime, timedelta, timezone

from dotenv import load_dotenv
from playwright.async_api import async_playwright

load_dotenv()

GOOGLE_EMAIL    = os.environ.get("GOOGLE_EMAIL", "")
GOOGLE_PASSWORD = os.environ.get("GOOGLE_PASSWORD", "")
LOOKER_URL      = "https://datastudio.google.com/u/0/reporting/67170489-51e9-43db-977a-ab7ec72bdb73/page/p_hhj93yd8xd"

SPREADSHEET_ID   = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"
GOOGLE_CREDS_ENV  = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE = "google_credentials.json"
SHEET_NAME        = "온라인매출"


def get_gspread_client():
    from google.oauth2.service_account import Credentials
    import gspread
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    if GOOGLE_CREDS_ENV:
        creds = Credentials.from_service_account_info(json.loads(GOOGLE_CREDS_ENV), scopes=scopes)
    else:
        creds = Credentials.from_service_account_file(GOOGLE_CREDS_FILE, scopes=scopes)
    return gspread.authorize(creds)


async def google_login(page):
    """Google 계정 로그인"""
    print("Google 로그인 중...")
    await page.goto("https://accounts.google.com/signin", wait_until="networkidle")
    await page.wait_for_timeout(2000)

    # 이메일 입력
    await page.fill('input[type="email"]', GOOGLE_EMAIL)
    await page.keyboard.press("Enter")
    await page.wait_for_timeout(2000)

    # 비밀번호 입력
    await page.fill('input[type="password"]', GOOGLE_PASSWORD)
    await page.keyboard.press("Enter")
    await page.wait_for_timeout(3000)

    # 추가 인증 화면 처리 (건너뛰기 버튼 있으면 클릭)
    for selector in ['button:has-text("건너뛰기")', 'button:has-text("Skip")',
                     'button:has-text("나중에")', 'button:has-text("Not now")']:
        try:
            btn = page.locator(selector).first
            if await btn.is_visible(timeout=2000):
                await btn.click()
                await page.wait_for_timeout(1000)
        except Exception:
            pass

    print(f"로그인 완료: {page.url}")


async def scrape_looker(page) -> list:
    """Looker Studio 대시보드에서 테이블 데이터 추출"""
    print(f"대시보드 접속 중: {LOOKER_URL}")
    await page.goto(LOOKER_URL, wait_until="networkidle", timeout=30000)
    await page.wait_for_timeout(5000)

    # 스크린샷 저장 (디버깅용)
    await page.screenshot(path="debug_looker.png", full_page=True)
    print("스크린샷 저장: debug_looker.png")

    # 테이블 데이터 추출 시도
    rows = []
    try:
        # Looker Studio 테이블 셀 선택자
        selectors = [
            "canvas",
            "[data-componenttype='table']",
            ".cell-container",
            "table tr",
        ]

        # 페이지 텍스트에서 날짜/숫자 패턴 추출
        content = await page.inner_text("body")
        lines = [l.strip() for l in content.split("\n") if l.strip()]

        print(f"\n--- 페이지 텍스트 (앞 80줄) ---")
        for line in lines[:80]:
            print(repr(line))

    except Exception as e:
        print(f"데이터 추출 오류: {e}")

    return rows


async def main():
    if not GOOGLE_EMAIL or not GOOGLE_PASSWORD:
        print("GOOGLE_EMAIL / GOOGLE_PASSWORD 환경변수 없음")
        return

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage"],
        )
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            locale="ko-KR",
            timezone_id="Asia/Seoul",
        )
        page = await context.new_page()

        try:
            await google_login(page)
            await page.screenshot(path="debug_after_login.png")

            rows = await scrape_looker(page)
            print(f"\n수집된 행: {len(rows)}")

        except Exception as e:
            print(f"오류: {e}")
            await page.screenshot(path="debug_error.png")
        finally:
            await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
