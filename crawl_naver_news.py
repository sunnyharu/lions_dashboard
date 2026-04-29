"""
네이버 뉴스/카페 검색 → Google Sheets '뉴스이슈' 탭 적재
어제 날짜 기준으로 '삼성 라이온즈' 관련 기사/카페글 수집
"""
import json
import os
import re
import requests
from datetime import datetime, timedelta

from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
import gspread

load_dotenv()

NAVER_CLIENT_ID     = os.environ.get("NAVER_CLIENT_ID", "")
NAVER_CLIENT_SECRET = os.environ.get("NAVER_CLIENT_SECRET", "")
SPREADSHEET_ID      = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"
SHEET_NAME          = "뉴스이슈"
GOOGLE_CREDS_ENV    = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE   = "google_credentials.json"

yesterday  = datetime.today() - timedelta(days=1)
DATE_STR   = yesterday.strftime("%Y.%m.%d")
DATE_YYYYMMDD = yesterday.strftime("%Y%m%d")

QUERY = "삼성 라이온즈"
MAX_DISPLAY = 10  # 탭당 최대 수집 건수


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


def strip_html(text: str) -> str:
    return re.sub(r"<[^>]+>", "", text).replace("&quot;", '"').replace("&amp;", "&").replace("&#39;", "'").strip()


def naver_search(endpoint: str, query: str, display: int = MAX_DISPLAY) -> list:
    headers = {
        "X-Naver-Client-Id":     NAVER_CLIENT_ID,
        "X-Naver-Client-Secret": NAVER_CLIENT_SECRET,
    }
    params = {"query": query, "display": display, "sort": "date"}
    resp = requests.get(
        f"https://openapi.naver.com/v1/search/{endpoint}.json",
        headers=headers, params=params, timeout=10,
    )
    if resp.status_code != 200:
        print(f"[{endpoint}] 오류: {resp.status_code} {resp.text[:200]}")
        return []
    return resp.json().get("items", [])


def filter_by_date(items: list, yyyymmdd: str, source: str) -> list:
    """pubDate 기준으로 어제 날짜 필터링 후 필요한 필드만 추출"""
    results = []
    for item in items:
        pub = item.get("pubDate", "")  # "Tue, 29 Apr 2026 10:00:00 +0900"
        try:
            pub_dt = datetime.strptime(pub, "%a, %d %b %Y %H:%M:%S %z")
            pub_yyyymmdd = pub_dt.strftime("%Y%m%d")
        except Exception:
            pub_yyyymmdd = ""

        if pub_yyyymmdd != yyyymmdd:
            continue

        title   = strip_html(item.get("title", ""))
        desc    = strip_html(item.get("description", ""))[:80]
        link    = item.get("link", "") or item.get("url", "")
        results.append([DATE_STR, source, title, desc, link])
    return results


def ensure_sheet(client):
    sh = client.open_by_key(SPREADSHEET_ID)
    try:
        ws = sh.worksheet(SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=1000, cols=5)
        ws.append_row(["날짜", "출처", "제목", "요약", "링크"])
        print(f"시트 '{SHEET_NAME}' 생성")
    return ws


def upload(ws, rows: list):
    if not rows:
        print("수집된 이슈 없음")
        return

    existing = ws.get_all_values()
    header   = existing[0] if existing else []
    date_col = header.index("날짜") if "날짜" in header else 0

    existing_keys = set()
    for r in existing[1:]:
        if len(r) > 2:
            existing_keys.add((r[date_col].strip(), r[2].strip()))  # (날짜, 제목)

    inserted = 0
    for row in rows:
        key = (row[0], row[2])
        if key in existing_keys:
            continue
        ws.append_row(row)
        existing_keys.add(key)
        inserted += 1

    print(f"뉴스이슈 적재: {inserted}건 삽입 (중복 {len(rows)-inserted}건 스킵)")


def main():
    if not NAVER_CLIENT_ID or not NAVER_CLIENT_SECRET:
        print("NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 환경변수가 없습니다. 스킵.")
        return

    print(f"네이버 검색 ({DATE_STR})...")
    news_items  = naver_search("news",         QUERY)
    cafe_items  = naver_search("cafearticle",  QUERY)

    rows  = filter_by_date(news_items,  DATE_YYYYMMDD, "뉴스")
    rows += filter_by_date(cafe_items,  DATE_YYYYMMDD, "카페")
    print(f"필터 후: 뉴스 {len([r for r in rows if r[1]=='뉴스'])}건 / 카페 {len([r for r in rows if r[1]=='카페'])}건")

    client = get_gspread_client()
    ws     = ensure_sheet(client)
    upload(ws, rows)


if __name__ == "__main__":
    main()
