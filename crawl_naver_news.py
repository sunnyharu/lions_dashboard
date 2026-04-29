"""
네이버 뉴스/카페 → Google Sheets 적재
- 뉴스: '삼성 라이온즈' 관련도순 최근 7일 상위 5건, Claude 요약
- 카페: 사자사랑방(lionsball) 키워드별 최근 7일 수집 → 조회수 상위 10건 + 트렌드 다이제스트
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
ANTHROPIC_API_KEY   = os.environ.get("ANTHROPIC_API_KEY", "")
SPREADSHEET_ID      = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"
GOOGLE_CREDS_ENV    = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE   = "google_credentials.json"

KST     = timezone(timedelta(hours=9))
NOW_KST = datetime.now(KST)

VALID_DATES = {(NOW_KST - timedelta(days=i)).strftime("%Y%m%d") for i in range(7)}
DATE_TODAY  = NOW_KST.strftime("%Y.%m.%d")

NEWS_QUERY    = "삼성 라이온즈"
NEWS_MAX      = 5
CAFE_TOP_N    = 10
CAFE_KEYWORDS = [
    "유니폼", "베리즈", "응원봉", "마킹키트", "로고볼",
    "짐색", "티셔츠", "백팩", "셔츠", "보스턴백",
    "볼캡", "자켓", "키링", "타월", "머플러", "어린이회원",
]
TRADE_KEYWORDS = ["판매", "팝니다", "팔아요", "삽니다", "구매", "거래", "양도", "나눔", "무료나눔", "중고", "원에"]

FETCH_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept-Language": "ko-KR,ko;q=0.9",
}

SHEET_NEWS    = "뉴스이슈"
SHEET_DIGEST  = "카페트렌드"
SHEET_HEADER  = ["날짜", "출처", "제목", "AI요약", "조회수", "링크"]


# ── Google Sheets ─────────────────────────────────────────────────────────────

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


def ensure_sheet(client, name: str, header: list):
    sh = client.open_by_key(SPREADSHEET_ID)
    try:
        ws = sh.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=name, rows=2000, cols=len(header))
        ws.append_row(header)
        print(f"시트 '{name}' 생성")
    return ws


# ── Claude ────────────────────────────────────────────────────────────────────

def call_claude(prompt: str, max_tokens: int = 300) -> str:
    if not ANTHROPIC_API_KEY:
        return ""
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=max_tokens,
            messages=[{"role": "user", "content": prompt}],
        )
        return msg.content[0].text.strip()
    except Exception as e:
        print(f"  Claude 오류: {e}")
        return ""


# ── 유틸 ──────────────────────────────────────────────────────────────────────

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
        return DATE_TODAY


def naver_search(endpoint: str, query: str, display: int = 100, sort: str = "sim") -> list:
    headers = {
        "X-Naver-Client-Id":     NAVER_CLIENT_ID,
        "X-Naver-Client-Secret": NAVER_CLIENT_SECRET,
    }
    resp = requests.get(
        f"https://openapi.naver.com/v1/search/{endpoint}.json",
        headers=headers,
        params={"query": query, "display": display, "sort": sort},
        timeout=10,
    )
    if resp.status_code != 200:
        print(f"  [{endpoint}] 오류: {resp.status_code}")
        return []
    return resp.json().get("items", [])


def fetch_text(url: str, max_chars: int = 3000) -> str:
    try:
        resp = requests.get(url, headers=FETCH_HEADERS, timeout=10)
        if resp.status_code != 200:
            return ""
        soup = BeautifulSoup(resp.text, "html.parser")
        for selector, attr in [
            ("div", {"id": "dic_area"}),
            ("div", {"class": "newsct_article"}),
            ("div", {"class": "_article_body"}),
            ("article", {}),
        ]:
            el = soup.find(selector, attr)
            if el:
                return el.get_text(" ", strip=True)[:max_chars]
        return soup.get_text(" ", strip=True)[:max_chars]
    except Exception:
        return ""


def fetch_cafe_view_count(url: str) -> int:
    try:
        resp = requests.get(url, headers=FETCH_HEADERS, timeout=10)
        if resp.status_code != 200:
            return 0
        soup = BeautifulSoup(resp.text, "html.parser")
        for sel in ["span.count_view", "em.count_view", "span.ViewCnt", ".view_count", "span.count_view_num"]:
            el = soup.select_one(sel)
            if el:
                nums = re.findall(r"\d+", el.text.replace(",", ""))
                if nums:
                    return int(nums[0])
        m = re.search(r"조회\s*[\D]{0,5}([\d,]+)", resp.text)
        if m:
            return int(m.group(1).replace(",", ""))
        return 0
    except Exception:
        return 0


def is_trade_post(title: str) -> bool:
    return any(kw in title for kw in TRADE_KEYWORDS)


# ── 뉴스 ──────────────────────────────────────────────────────────────────────

def process_news() -> list:
    print(f"\n[뉴스] '{NEWS_QUERY}' 관련도순 최근 7일 상위 {NEWS_MAX}건")
    items   = naver_search("news", NEWS_QUERY, display=50, sort="sim")
    results = []
    for item in items:
        if len(results) >= NEWS_MAX:
            break
        pub_date = parse_pub_date(item.get("pubDate", ""))
        if pub_date not in VALID_DATES:
            continue
        title = strip_html(item.get("title", ""))
        link  = item.get("link", "") or item.get("originallink", "")
        desc  = strip_html(item.get("description", ""))

        body    = fetch_text(link)
        summary = call_claude(
            f"다음 기사를 한 줄로 요약해줘. 상품 판매·출시 정보가 있으면 꼭 포함해.\n\n{(body or desc)[:2000]}"
        ) or desc[:120]

        print(f"  → {yyyymmdd_to_str(pub_date)} | {title[:35]}")
        results.append({
            "date":    yyyymmdd_to_str(pub_date),
            "source":  "뉴스",
            "title":   title,
            "summary": summary,
            "views":   0,
            "link":    link,
        })
        time.sleep(0.3)
    return results


# ── 카페 ──────────────────────────────────────────────────────────────────────

def collect_cafe_posts() -> list:
    print(f"\n[카페] 키워드별 수집 (최근 7일, lionsball)")
    seen   = set()
    posts  = []
    first_item_logged = False
    for kw in CAFE_KEYWORDS:
        items = naver_search("cafearticle", f"삼성 라이온즈 {kw}", display=30, sort="sim")
        if items and not first_item_logged:
            print(f"  [디버그] 첫 카페 항목 키: {list(items[0].keys())}")
            print(f"  [디버그] link={items[0].get('link','')[:60]}")
            print(f"  [디버그] cafeurl={items[0].get('cafeurl','')[:60]}")
            print(f"  [디버그] cafename={items[0].get('cafename','')}")
            first_item_logged = True
        for item in items:
            link     = item.get("link", "") or item.get("url", "")
            cafeurl  = item.get("cafeurl", "")
            cafename = item.get("cafename", "")
            is_lionsball = (
                "lionsball" in link or
                "lionsball" in cafeurl or
                "사자사랑방" in cafename.replace(" ", "")
            )
            if not is_lionsball or link in seen:
                continue
            pub_date = parse_pub_date(item.get("pubDate", ""))
            # 파싱 실패 시 수집 허용 (sort=sim이므로 최근 글일 가능성 높음)
            if pub_date and pub_date not in VALID_DATES:
                continue
            title = strip_html(item.get("title", ""))
            if is_trade_post(title):
                continue
            seen.add(link)
            posts.append({
                "date":    yyyymmdd_to_str(pub_date) if pub_date else DATE_TODAY,
                "title":   title,
                "desc":    strip_html(item.get("description", ""))[:200],
                "link":    link,
                "keyword": kw,
                "views":   0,
            })

    print(f"  중복제거 후 {len(posts)}건")

    print("  조회수 추출 중...")
    for p in posts:
        p["views"] = fetch_cafe_view_count(p["link"])
        time.sleep(0.2)

    return posts


def process_cafe_top10(posts: list) -> list:
    top10   = sorted(posts, key=lambda x: x["views"], reverse=True)[:CAFE_TOP_N]
    print(f"\n[카페 TOP10] Claude 요약")
    results = []
    for p in top10:
        body    = fetch_text(p["link"])
        summary = call_claude(
            f"다음 카페 글의 주요 내용을 한 줄로 요약해줘.\n제목: {p['title']}\n본문: {(body or p['desc'])[:1500]}"
        ) or p["desc"][:120]
        print(f"  조회수 {p['views']:,} | {p['title'][:30]}")
        results.append({
            "date":    p["date"],
            "source":  "카페(사자사랑방)",
            "title":   p["title"],
            "summary": summary,
            "views":   p["views"],
            "link":    p["link"],
        })
        time.sleep(0.3)
    return results


def generate_cafe_digest(posts: list) -> str:
    if not posts:
        return ""
    post_text = "\n".join(
        f"- [{p['keyword']}] {p['title']}: {p['desc'][:80]}"
        for p in posts[:50]
    )
    digest = call_claude(
        f"다음은 삼성 라이온즈 팬 카페(사자사랑방)의 최근 7일간 글 목록입니다.\n"
        f"팬들이 주로 관심 갖는 상품, 이슈, 반응 등을 3~5줄로 정리해줘.\n\n{post_text}",
        max_tokens=500,
    )
    print(f"\n[카페 다이제스트] 생성 완료")
    return digest


# ── 업로드 ────────────────────────────────────────────────────────────────────

def upload_rows(ws, rows: list):
    if not rows:
        print("  업로드할 항목 없음")
        return
    existing  = ws.get_all_values()
    header    = existing[0] if existing else SHEET_HEADER
    date_col  = header.index("날짜")  if "날짜"  in header else 0
    title_col = header.index("제목")  if "제목"  in header else 2

    existing_keys = {
        (r[date_col].strip(), r[title_col].strip())
        for r in existing[1:] if len(r) > title_col
    }
    new_rows = []
    for r in rows:
        key = (r["date"], r["title"])
        if key in existing_keys:
            continue
        new_rows.append([r["date"], r["source"], r["title"], r["summary"], r.get("views", 0), r["link"]])
        existing_keys.add(key)

    if new_rows:
        ws.append_rows(new_rows, value_input_option="USER_ENTERED")
    print(f"  {len(new_rows)}건 삽입 (중복 {len(rows) - len(new_rows)}건 스킵)")


def upload_digest(ws, digest: str):
    if not digest:
        return
    existing_dates = {r[0].strip() for r in ws.get_all_values()[1:] if r}
    if DATE_TODAY not in existing_dates:
        ws.append_row([DATE_TODAY, digest])
        print(f"  카페트렌드 저장: {DATE_TODAY}")
    else:
        print(f"  카페트렌드 이미 존재: {DATE_TODAY}")


# ── 메인 ──────────────────────────────────────────────────────────────────────

def main():
    if not NAVER_CLIENT_ID or not NAVER_CLIENT_SECRET:
        print("NAVER 환경변수 없음. 스킵.")
        return

    news_rows  = process_news()
    cafe_posts = collect_cafe_posts()
    cafe_top10 = process_cafe_top10(cafe_posts)
    digest     = generate_cafe_digest(cafe_posts)

    client    = get_gspread_client()
    ws_news   = ensure_sheet(client, SHEET_NEWS,   SHEET_HEADER)
    ws_digest = ensure_sheet(client, SHEET_DIGEST, ["날짜", "다이제스트"])

    print("\n[업로드] 뉴스이슈")
    upload_rows(ws_news, news_rows + cafe_top10)
    print("[업로드] 카페트렌드")
    upload_digest(ws_digest, digest)


if __name__ == "__main__":
    main()
