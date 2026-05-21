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
        ws = sh.add_worksheet(title=SHEET_NAME, rows=50000, cols=len(SHEET_HEADER))
        ws.append_row(SHEET_HEADER)
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

SHEET_HEADER = ["판매일자", "상품코드", "상품명", "칼라명", "사이즈명", "자사바코드", "판매단가", "판매수량", "실판매금액"]

# 집계 기준 컬럼 (이 조합이 같으면 수량/금액 합산)
GROUP_COLS  = ["상품코드", "상품명", "칼라명", "사이즈명", "자사바코드", "판매단가"]
SUM_COLS    = ["판매수량", "실판매금액"]


def parse_and_aggregate(excel_bytes: bytes) -> list:
    """엑셀 파일 읽어서 상품별로 집계"""
    df = pd.read_excel(io.BytesIO(excel_bytes), engine="openpyxl")
    df.columns = df.columns.str.strip()
    print(f"엑셀 로드: {len(df)}행, 컬럼: {list(df.columns)}")

    # 필요한 컬럼만 추출 (없는 컬럼은 빈값으로)
    for col in GROUP_COLS + SUM_COLS:
        if col not in df.columns:
            print(f"  컬럼 없음 (빈값 처리): {col}")
            df[col] = "" if col in GROUP_COLS else 0

    # 숫자 컬럼 변환
    for col in SUM_COLS:
        df[col] = (
            df[col].astype(str)
            .str.replace(",", "").str.strip()
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # 판매단가도 숫자 변환
    df["판매단가"] = (
        df["판매단가"].astype(str)
        .str.replace(",", "").str.strip()
    )
    df["판매단가"] = pd.to_numeric(df["판매단가"], errors="coerce").fillna(0)

    # 집계
    agg = (
        df.groupby(GROUP_COLS, as_index=False)
        .agg({col: "sum" for col in SUM_COLS})
    )
    agg = agg.sort_values("실판매금액", ascending=False)
    print(f"집계 결과: {len(agg)}개 상품 SKU")

    rows = []
    for _, r in agg.iterrows():
        rows.append([
            DATE_SHEET,
            str(r["상품코드"]).strip(),
            str(r["상품명"]).strip(),
            str(r["칼라명"]).strip(),
            str(r["사이즈명"]).strip(),
            str(r["자사바코드"]).strip(),
            int(r["판매단가"]),
            int(r["판매수량"]),
            int(r["실판매금액"]),
        ])
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

        menu_item = page.locator("text=매장판매일보").first
        await menu_item.wait_for(timeout=5000)
        await menu_item.click()
        await page.wait_for_timeout(4000)
        await page.screenshot(path="debug_02_menu.png")
        print(f"메뉴 이동 완료: {page.url}")

        # ── iframe 확인 → 컨텍스트 선택 ────────────
        import re as _re
        frames = page.frames
        print(f"프레임 수: {len(frames)}")
        for f in frames:
            print(f"  frame: url={f.url}")

        # 매장판매일보 내용이 있는 frame 찾기 (날짜 input 기준)
        target_frame = page  # 기본은 main page
        for f in frames:
            try:
                inputs = await f.locator('input[type="text"]').all()
                for inp in inputs:
                    val = await inp.input_value()
                    if _re.match(r'\d{4}-\d{2}-\d{2}', val.strip()):
                        print(f"  날짜 input 있는 frame: {f.url}")
                        target_frame = f
                        break
            except Exception:
                continue

        # ── 날짜 설정 ────────────────────────────────
        print(f"날짜 설정: {DATE_PARAM}")
        date_inputs = await target_frame.locator('input[type="text"]').all()
        set_count = 0
        for inp in date_inputs[:10]:
            try:
                val = await inp.input_value()
                print(f"  input 발견: '{val}'")
                if _re.match(r'\d{4}-\d{2}-\d{2}', val.strip()):
                    await inp.click()
                    await page.keyboard.press("Control+a")
                    await page.keyboard.type(DATE_PARAM)
                    await page.keyboard.press("Tab")
                    await page.wait_for_timeout(400)
                    new_val = await inp.input_value()
                    print(f"  → 변경 후: '{new_val}'")
                    set_count += 1
            except Exception as e:
                print(f"  input 처리 오류: {e}")
        print(f"날짜 입력 필드 {set_count}개 설정")
        await page.screenshot(path="debug_03_date.png")

        # ── 조회 클릭 ────────────────────────────────
        print("조회 클릭...")
        조회_btn = target_frame.locator("a:has-text('조회'), button:has-text('조회'), [title='조회']").first
        await 조회_btn.wait_for(timeout=10000)
        await 조회_btn.click()
        await page.wait_for_timeout(10000)
        await page.screenshot(path="debug_04_result.png")
        print("조회 완료")

        # ── 모달 처리 함수 (iframe 대상) ────────────────
        async def dismiss_modal(frame):
            """iframe 안의 모달 확인 버튼 클릭"""
            modal_sel = ".modal.in button, .modal.fade.in button"
            try:
                btns = await frame.locator(modal_sel).all()
                print(f"  모달 버튼 수: {len(btns)}")
                for btn in btns:
                    txt = (await btn.text_content() or "").strip()
                    print(f"  모달 버튼: '{txt}'")
                if btns:
                    await btns[0].click()
                    await page.wait_for_timeout(1000)
                    txt = (await btns[0].text_content() or "").strip()
                    print(f"  모달 클릭: '{txt}'")
                    return True
            except Exception as e:
                print(f"  모달 처리 오류: {e}")
            return False

        # ── 엑셀 다운로드 ────────────────────────────
        print("엑셀 다운로드 중...")
        excel_bytes = None
        excel_btn   = target_frame.locator("a:has-text('엑셀'), button:has-text('엑셀'), [title='엑셀']").first

        # 엑셀 클릭 전 기존 모달 정리
        await dismiss_modal(target_frame)

        try:
            async with page.expect_download(timeout=25000) as dl_info:
                await excel_btn.click()
                # 클릭 후 모달 뜨면 확인 처리
                await page.wait_for_timeout(2000)
                await dismiss_modal(target_frame)

            download    = await dl_info.value
            excel_bytes = await download.read()
            print(f"엑셀 다운로드 완료: {len(excel_bytes):,} bytes")
        except Exception as e1:
            print(f"다운로드 실패: {e1}")
            await page.screenshot(path="debug_05_excel_error.png")

        if not excel_bytes:
            print("엑셀 다운로드 실패")
            await browser.close()
            return None

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
