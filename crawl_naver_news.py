"""
네이버 뉴스/카페 검색 → Google Sheets '뉴스이슈' 탭 적재
- 뉴스: '삼성 라이온즈 베리즈' 키워드
- 카페: '삼성 라이온즈 + 상품 키워드' 다중 검색, 거래글 제외
"""
import json
import os
import re
import time
import requests
from datetime import datetime, timedelta, timezone
from bs4 import BeautifulSoup

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

KST       = timezone(timedelta(hours=9))
now_kst   = datetime.now(KST)
yesterday = now_kst - timedelta(days=1)

DATE_STR      = yesterday.strftime("%Y.%m.%d")
DATE_YYYYMMDD = yesterday.strftime("%Y%m%d")
VALID_DATES   = {DATE_YYYYMMDD, now_kst.strftime("%Y%m%d")}  # 어제 + 오늘

NEWS_QUERY    = "삼성라이온즈 베리즈"
NEWS_MAX_DAY  = 5   # 하루 최대 뉴스 건수

CAFE_KEYWORDS = [
    "유니폼", "베리즈", "응원봉", "마킹키트", "로고볼",
    "짐색", "티셔츠", "백팩", "셔츠", "보스턴백",
    "볼캡", "자켓", "키링", "타월", "머플러", "어린이회원",
]

MAX_DISPLAY  = 100   # 네이버 API 최대값
TARGET_CAFE  = "사자사랑방"   # 공백 포함 변형도 허용: "사자 사랑방"

TRADE_KEYWORDS = ["판매", "팝니다", "팔아요", "삽니다", "구매", "거래", "양도", "나눔", "무료나눔", "중고", "원에"]


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
    for fmt in ("%a, %d %b %Y %H:%M:%S %z", "%Y-%m-%dT%H:%M:%S%z"):
        try:
            return datetime.strptime(pub, fmt).astimezone(KST).strftime("%Y%m%d")
        except Exception:
            pass
    return ""


def yyyymmdd_to_str(d: str) -> str:
    try:
        return f"{d[:4]}.{d[4:6]}.{d[6:8]}"
    except Exception:
        return DATE_STR


def naver_search(endpoint: str, query: str, display: int = MAX_DISPLAY, sort: str = "date") -> list:
    headers = {
        "X-Naver-Client-Id":     NAVER_CLIENT_ID,
        "X-Naver-Client-Secret": NAVER_CLIENT_SECRET,
    }
    params = {"query": query, "display": display, "sort": sort}
    resp = requests.get(
        f"https://openapi.naver.com/v1/search/{endpoint}.json",
        headers=headers, params=params, timeout=10,
    )
    if resp.status_code != 200:
        print(f"  [{endpoint}] 오류: {resp.status_code} {resp.text[:200]}")
        return []
    return resp.json().get("items", [])


def is_trade_post(title: str) -> bool:
    return any(kw in title for kw in TRADE_KEYWORDS)


NEWS_REQUIRED = ["삼성라이온즈", "베리즈"]

ARTICLE_SELECTORS = [
    "div#dic_area",           # 네이버 뉴스 본문
    "div.newsct_article",
    "div._article_body",
    "div.article_body",
    "article",
    "div#content",
]

FETCH_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept-Language": "ko-KR,ko;q=0.9",
}


def fetch_article_text(url: str) -> str:
    """기사 URL 본문 텍스트 반환. 실패 시 빈 문자열."""
    try:
        resp = requests.get(url, headers=FETCH_HEADERS, timeout=10)
        if resp.status_code != 200:
            return ""
        soup = BeautifulSoup(resp.text, "html.parser")
        for sel in ARTICLE_SELECTORS:
            tag, _, cls = sel.partition(".")
            attr = {"class": cls} if cls else {}
            if "#" in sel:
                tag, _, id_ = sel.partition("#")
                attr = {"id": id_}
            el = soup.find(tag or True, attr)
            if el:
                return el.get_text(" ", strip=True)
        return soup.get_text(" ", strip=True)[:3000]
    except Exception as e:
        print(f"    본문 fetch 오류: {e}")
        return ""


def has_all_keywords(text: str) -> bool:
    normalized = text.replace(" ", "")
    return all(kw.replace(" ", "") in normalized for kw in NEWS_REQUIRED)


def process_news(items: list) -> list:
    results = []
    for item in items:
        if len(results) >= NEWS_MAX_DAY:
            break
        raw_pub  = item.get("pubDate", "")
        pub_date = parse_pub_date(raw_pub)
        if pub_date not in VALID_DATES:
            continue
        title = strip_html(item.get("title", ""))
        link  = item.get("link", "") or item.get("originallink", "")
        print(f"  [뉴스] 본문 확인 중: {title[:30]}")
        body = fetch_article_text(link)
        if not has_all_keywords(title + body):
            print(f"    → 키워드 미충족 스킵")
            continue
        desc = strip_html(item.get("description", ""))[:100]
        print(f"    → 수집 ({pub_date})")
        results.append([yyyymmdd_to_str(pub_date), "뉴스", title, desc, link])
        time.sleep(0.5)
    return results


def is_target_cafe(cafe_name: str) -> bool:
    normalized = cafe_name.replace(" ", "")
    return TARGET_CAFE.replace(" ", "") in normalized


def process_cafe(items: list, keyword: str) -> list:
    results = []
    for item in items:
        cafe_name = strip_html(item.get("cafename", ""))
        if not is_target_cafe(cafe_name):
            print(f"  [카페/{keyword}] 제외(카페명: {cafe_name!r})")
            continue
        raw_pub  = item.get("pubDate", "")
        pub_date = parse_pub_date(raw_pub)
        date_str = yyyymmdd_to_str(pub_date) if pub_date in VALID_DATES else DATE_STR
        title    = strip_html(item.get("title", ""))
        print(f"  [카페/{keyword}] {pub_date or '?'} | {title[:30]}")
        if is_trade_post(title):
            print(f"    → 거래글 제외")
            continue
        desc = strip_html(item.get("description", ""))[:100]
        link = item.get("link", "") or item.get("url", "")
        results.append([date_str, "카페(사자사랑방)", title, desc, link])
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

    new_rows = []
    for row in rows:
        key = (row[0], row[2])
        if key in existing_keys:
            continue
        new_rows.append(row)
        existing_keys.add(key)

    if new_rows:
        ws.append_rows(new_rows, value_input_option="USER_ENTERED")

    print(f"뉴스이슈 적재: {len(new_rows)}건 삽입 (중복 {len(rows)-len(new_rows)}건 스킵)")


def main():
    if not NAVER_CLIENT_ID or not NAVER_CLIENT_SECRET:
        print("NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 환경변수가 없습니다. 스킵.")
        return

    print(f"=== 네이버 검색 ({DATE_STR} KST) ===")

    # 뉴스: 삼성라이온즈 베리즈 (관련도순, 하루 최대 5건)
    print(f"\n[뉴스] 쿼리: {NEWS_QUERY!r}")
    news_items = naver_search("news", NEWS_QUERY, display=50, sort="sim")
    rows = process_news(news_items)
    print(f"  → {len(rows)}건 수집")

    # 카페: 삼성 라이온즈 + 각 키워드
    cafe_rows = []
    seen_titles = set()
    for kw in CAFE_KEYWORDS:
        query = f"삼성 라이온즈 {kw}"
        print(f"\n[카페] 쿼리: {query!r}")
        items  = naver_search("cafearticle", query, display=MAX_DISPLAY)
        for r in process_cafe(items, kw):
            if r[2] not in seen_titles:  # 제목 기준 중복 제거
                cafe_rows.append(r)
                seen_titles.add(r[2])
    print(f"\n카페 총 {len(cafe_rows)}건 수집")
    rows += cafe_rows

    client = get_gspread_client()
    ws     = ensure_sheet(client)
    upload(ws, rows)


if __name__ == "__main__":
    main()
