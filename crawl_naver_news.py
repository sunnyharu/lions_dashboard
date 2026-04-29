"""
네이버 뉴스/카페 검색 → Google Sheets '뉴스이슈' 탭 적재
어제 날짜 기준으로 '삼성 라이온즈' 관련 기사/카페글 수집
카페글은 중고거래성 글(중고나라, 판매/구매 등) 제외
"""
import json
import os
import re
import requests
from datetime import datetime, timedelta, timezone

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

KST = timezone(timedelta(hours=9))
now_kst   = datetime.now(KST)
yesterday = now_kst - timedelta(days=1)

DATE_STR      = yesterday.strftime("%Y.%m.%d")
DATE_YYYYMMDD = yesterday.strftime("%Y%m%d")

QUERY       = "삼성 라이온즈"
MAX_DISPLAY = 20

# 거래글 제외 키워드 (제목 포함 시 스킵)
TRADE_KEYWORDS = ["판매", "팝니다", "팔아요", "삽니다", "구매", "거래", "양도", "나눔", "무료나눔", "중고", "원에"]

# 거래성 카페명
TRADE_CAFES = ["중고나라", "번개장터", "당근마켓", "클리앙중고장터", "중고장터"]


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


def parse_pub_date(pub: str) -> str:
    """pubDate → 'YYYYMMDD' (KST 기준). 실패 시 빈 문자열."""
    for fmt in ("%a, %d %b %Y %H:%M:%S %z", "%Y-%m-%dT%H:%M:%S%z"):
        try:
            dt = datetime.strptime(pub, fmt).astimezone(KST)
            return dt.strftime("%Y%m%d")
        except Exception:
            pass
    return ""


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


def is_trade_post(title: str, cafe_name: str = "") -> bool:
    """거래글 여부 판별"""
    if any(tc in cafe_name for tc in TRADE_CAFES):
        return True
    return any(kw in title for kw in TRADE_KEYWORDS)


def process_news(items: list) -> list:
    """뉴스 아이템 → [날짜, 출처, 제목, 요약, 링크] 리스트 (어제 날짜만)"""
    results = []
    for item in items:
        pub_date = parse_pub_date(item.get("pubDate", ""))
        if pub_date != DATE_YYYYMMDD:
            continue
        title = strip_html(item.get("title", ""))
        desc  = strip_html(item.get("description", ""))[:100]
        link  = item.get("link", "") or item.get("originallink", "")
        results.append([DATE_STR, "뉴스", title, desc, link])
    return results


def process_cafe(items: list) -> list:
    """카페 아이템 → [날짜, 출처, 제목, 요약, 링크] 리스트 (어제 날짜 + 거래글 제외)"""
    results = []
    for item in items:
        pub_date  = parse_pub_date(item.get("pubDate", ""))
        if pub_date != DATE_YYYYMMDD:
            continue
        title     = strip_html(item.get("title", ""))
        cafe_name = strip_html(item.get("cafename", ""))
        if is_trade_post(title, cafe_name):
            print(f"  거래글 제외: [{cafe_name}] {title[:30]}")
            continue
        desc = strip_html(item.get("description", ""))[:100]
        link = item.get("link", "") or item.get("url", "")
        results.append([DATE_STR, f"카페({cafe_name})" if cafe_name else "카페", title, desc, link])
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

    existing_keys = {
        (r[date_col].strip(), r[2].strip())
        for r in existing[1:]
        if len(r) > 2
    }

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

    print(f"네이버 검색 ({DATE_STR}, KST)...")
    news_items = naver_search("news",        QUERY)
    cafe_items = naver_search("cafearticle", QUERY)
    print(f"  뉴스 원본 {len(news_items)}건 / 카페 원본 {len(cafe_items)}건")

    rows  = process_news(news_items)
    rows += process_cafe(cafe_items)
    print(f"  날짜 필터 후: 뉴스 {len([r for r in rows if r[1]=='뉴스'])}건 / 카페 {sum(1 for r in rows if '카페' in r[1])}건")

    client = get_gspread_client()
    ws     = ensure_sheet(client)
    upload(ws, rows)


if __name__ == "__main__":
    main()
