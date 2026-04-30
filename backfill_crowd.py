"""
KBO 관중수 페이지에서 삼성 홈 경기 관중수를 소급 입력
https://www.koreabaseball.com/Record/Crowd/GraphDaily.aspx
"""
import json
import os
import re
import time

import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
import gspread

load_dotenv()

SPREADSHEET_ID    = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"
SHEET_NAME        = "경기현황"
GOOGLE_CREDS_ENV  = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE = "google_credentials.json"
SEASON            = "2026"
MONTHS            = ["03", "04", "05"]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept-Language": "ko-KR,ko;q=0.9",
}


def get_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    if GOOGLE_CREDS_ENV:
        creds = Credentials.from_service_account_info(json.loads(GOOGLE_CREDS_ENV), scopes=scopes)
    else:
        creds = Credentials.from_service_account_file(GOOGLE_CREDS_FILE, scopes=scopes)
    return gspread.authorize(creds)


def fetch_crowd_map() -> dict:
    """KBO 관중수 페이지에서 삼성 홈 경기 관중수 수집 → {YYYY.MM.DD: 관중수}"""
    url = "https://www.koreabaseball.com/Record/Crowd/GraphDaily.aspx"
    crowd_map = {}

    for month in MONTHS:
        params = {
            "season": SEASON,
            "month":  month,
            "team":   "삼성",
            "homeAway": "홈",
        }
        try:
            resp = requests.get(url, headers=HEADERS, params=params, timeout=15)
            resp.raise_for_status()
        except Exception as e:
            print(f"  {month}월 요청 오류: {e}")
            continue

        soup  = BeautifulSoup(resp.text, "html.parser")
        table = soup.find("table")
        if not table:
            print(f"  {month}월 테이블 없음")
            continue

        for row in table.find_all("tr")[1:]:  # 헤더 제외
            cols = [td.get_text(strip=True) for td in row.find_all("td")]
            if len(cols) < 6:
                continue
            # cols: [날짜, 요일, 홈, 방문, 구장, 관중수]
            date_raw = cols[0]   # "2026/03/28"
            home     = cols[2]   # "삼성"
            crowd_str = cols[5]  # "24,000"

            if home != "삼성":
                continue

            # 날짜 포맷 변환: "2026/03/28" → "2026.03.28"
            date_fmt = date_raw.replace("/", ".")
            crowd    = int(crowd_str.replace(",", "")) if crowd_str.replace(",", "").isdigit() else 0

            crowd_map[date_fmt] = crowd
            print(f"  {date_fmt} 관중수: {crowd:,}")

        time.sleep(0.3)

    return crowd_map


def main():
    print("KBO 관중수 수집 중...")
    crowd_map = fetch_crowd_map()
    print(f"\n총 {len(crowd_map)}건 수집\n")

    if not crowd_map:
        print("수집된 데이터 없음")
        return

    client   = get_client()
    sh       = client.open_by_key(SPREADSHEET_ID)
    ws       = sh.worksheet(SHEET_NAME)
    all_vals = ws.get_all_values()

    if not all_vals:
        print("시트 비어있음")
        return

    header = all_vals[0]

    # 관중수 컬럼 없으면 추가
    if "관중수" not in header:
        header.append("관중수")
        ws.update([header], f"A1:{chr(65 + len(header) - 1)}1")
        print("헤더에 '관중수' 컬럼 추가")

    col_date  = header.index("날짜")
    col_ha    = header.index("홈/어웨이")
    col_crowd = header.index("관중수")

    updated = 0
    for i, row in enumerate(all_vals[1:], start=2):
        if len(row) <= col_date:
            continue
        date_val = row[col_date].strip()   # "2026.04.01"
        ha       = row[col_ha].strip() if len(row) > col_ha else ""

        if ha != "홈":
            continue

        # 기존 관중수가 있으면 스킵
        existing = row[col_crowd].strip() if len(row) > col_crowd else ""
        if existing and existing not in ("0", "-", ""):
            print(f"  이미 있음 스킵: {date_val} ({existing})")
            continue

        crowd = crowd_map.get(date_val, 0)
        if not crowd:
            print(f"  관중수 없음: {date_val}")
            continue

        cell_addr = f"{chr(65 + col_crowd)}{i}"
        ws.update([[crowd]], cell_addr)
        print(f"  ✓ {date_val} 관중수 입력: {crowd:,}")
        updated += 1
        time.sleep(0.1)

    print(f"\n완료: {updated}건 업데이트")


if __name__ == "__main__":
    main()
