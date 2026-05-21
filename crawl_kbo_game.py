"""
KBO 공식 API에서 삼성 라이온즈 어제 경기 결과를 가져와
Google Sheets '경기현황' 탭에 적재
"""
import json
import os
import re
from datetime import datetime, timedelta

import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
import gspread

load_dotenv()

# ── Google Sheets 설정 ──────────────────────────────────
SPREADSHEET_ID    = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"
SHEET_NAME        = "경기현황"
GOOGLE_CREDS_ENV  = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE = "google_credentials.json"

# ── 조회 기준: 어제 ──────────────────────────────────────
yesterday  = datetime.today() - timedelta(days=1)
SEASON     = str(yesterday.year)
GAME_MONTH = yesterday.strftime("%m")        # "04"
DAY_TEXT   = yesterday.strftime("%m.%d")     # "04.28"
DATE_STR   = yesterday.strftime("%Y.%m.%d")  # "2026.04.28"

STADIUM_CAPACITY = 24000  # 대구 삼성 라이온즈 파크 수용인원
KBO_API = "https://www.koreabaseball.com/ws/Schedule.asmx/GetScheduleList"
SHEET_HEADER = ["날짜", "홈/어웨이", "상대팀", "결과", "관중수"]


KBO_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Accept-Language": "ko-KR,ko;q=0.9",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "X-Requested-With": "XMLHttpRequest",
    "Referer": "https://www.koreabaseball.com/Schedule/Schedule.aspx",
    "Origin": "https://www.koreabaseball.com",
}


def fetch_kbo_game() -> dict | None:
    """어제 삼성 라이온즈 경기 정보를 KBO API에서 조회"""
    payload = {
        "leId":      "1",
        "srIdList":  "0,9",
        "seasonId":  SEASON,
        "startDay":  yesterday.strftime("%Y%m01"),
        "endDay":    yesterday.strftime("%Y%m%d"),
        "gameMonth": GAME_MONTH,
        "teamId":    "SS",
    }
    resp = requests.post(KBO_API, headers=KBO_HEADERS, data=payload, timeout=15)
    resp.raise_for_status()

    if not resp.text.strip():
        print(f"  KBO API 빈 응답 (경기 없음 또는 차단)")
        return None

    try:
        data = resp.json()
    except Exception as e:
        print(f"  KBO API JSON 파싱 오류: {e}\n  응답 내용: {resp.text[:200]}")
        return None

    rows = data.get("rows", [])
    for entry in rows:
        cells = entry.get("row", [])
        if not cells:
            continue

        # 날짜 셀 (class="day")
        day_cell = next((c for c in cells if c.get("Class") == "day"), None)
        if not day_cell:
            continue
        day_raw = re.sub(r"\(.*?\)", "", day_cell.get("Text", "")).strip()  # "04.28"
        if day_raw != DAY_TEXT:
            continue

        # 경기 셀 (class="play")
        play_cell = next((c for c in cells if c.get("Class") == "play"), None)
        if not play_cell:
            continue

        game = parse_play(play_cell["Text"])
        if game:
            # 홈 경기일 때만 관중수 수집
            crowd = 0
            if game.get("홈/어웨이") == "홈":
                crowd = fetch_crowd()
            game["관중수"] = crowd
        return game

    print(f"어제({DAY_TEXT}) 삼성 라이온즈 경기 없음")
    return None


def fetch_crowd() -> int:
    """KBO 관중수 페이지에서 어제 삼성 홈 경기 관중수 조회"""
    url = "https://www.koreabaseball.com/Record/Crowd/GraphDaily.aspx"
    params = {"season": SEASON, "month": GAME_MONTH, "team": "삼성", "homeAway": "홈"}
    try:
        resp = requests.get(url, headers=KBO_HEADERS, params=params, timeout=15)
        resp.raise_for_status()
        print(f"  관중수 페이지 응답: {resp.status_code} ({len(resp.text)}자)")
        soup  = BeautifulSoup(resp.text, "html.parser")
        table = soup.find("table")
        if not table:
            print("  관중수 테이블 없음")
            return 0
        # DATE_STR = "2026.04.28" → "2026/04/28"
        target_date = DATE_STR.replace(".", "/")
        print(f"  관중수 조회 날짜: {target_date}")
        for row in table.find_all("tr")[1:]:
            cols = [td.get_text(strip=True) for td in row.find_all("td")]
            print(f"  관중수 테이블 행: {cols}")  # 컬럼 구조 확인용
            if len(cols) < 4:
                continue
            # 날짜 매칭 (cols[0])
            if cols[0] == target_date:
                # 관중수는 마지막 숫자 컬럼
                for col in reversed(cols):
                    crowd_str = col.replace(",", "").replace(" ", "")
                    if crowd_str.isdigit() and int(crowd_str) > 1000:
                        print(f"  관중수 발견: {col}")
                        return int(crowd_str)
        print(f"  날짜 {target_date} 관중수 행 없음")
    except Exception as e:
        print(f"  관중수 조회 오류: {e}")
    return 0


def parse_play(html: str) -> dict:
    """
    play 셀 HTML 파싱
    포맷: <span>원정팀</span><em><span class="win/lose">점수</span>vs<span class="win/lose">점수</span></em><span>홈팀</span>
    왼쪽=원정, 오른쪽=홈
    """
    soup = BeautifulSoup(html, "html.parser")
    spans = soup.find_all("span", recursive=False)

    if len(spans) < 2:
        return {}

    away_team = spans[0].get_text(strip=True)
    home_team = spans[-1].get_text(strip=True)

    # em 안의 숫자 span만 추출 → [원정점수span, 홈점수span]
    score_elems = [
        s for s in soup.select("em span")
        if re.fullmatch(r"\d+", s.get_text(strip=True))
    ]
    away_score_cls = score_elems[0].get("class", [""])[0] if len(score_elems) >= 1 else ""
    home_score_cls = score_elems[1].get("class", [""])[0] if len(score_elems) >= 2 else ""

    # 삼성 관점
    if home_team == "삼성":
        home_away          = "홈"
        opponent           = away_team
        samsung_result_cls = home_score_cls
    else:
        home_away          = "어웨이"
        opponent           = home_team
        samsung_result_cls = away_score_cls

    result_map = {"win": "승", "lose": "패", "draw": "무"}
    result = result_map.get(samsung_result_cls, "-")

    return {
        "날짜":     DATE_STR,
        "홈/어웨이": home_away,
        "상대팀":   opponent,
        "결과":     result,
    }


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


def upload_to_sheets(game: dict):
    client = get_gspread_client()
    sh     = client.open_by_key(SPREADSHEET_ID)

    try:
        ws = sh.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=500, cols=10)

    existing = ws.get_all_values()
    if not existing:
        ws.append_row(SHEET_HEADER)
        existing = [SHEET_HEADER]

    # 헤더에 관중수 없으면 추가
    header = existing[0]
    if "관중수" not in header:
        ws.update([SHEET_HEADER], "A1:E1")
        print("헤더에 '관중수' 컬럼 추가")
        header = SHEET_HEADER

    # 중복 날짜 체크
    date_col = header.index("날짜") if "날짜" in header else 0
    existing_dates = {r[date_col].strip() for r in existing[1:] if r}
    if game["날짜"] in existing_dates:
        print(f"이미 존재하는 날짜 스킵: {game['날짜']}")
        return

    row = [game["날짜"], game["홈/어웨이"], game["상대팀"], game["결과"], game.get("관중수", 0)]
    ws.append_row(row)
    print(f"Google Sheets 적재 완료: {row}")


def main():
    game = fetch_kbo_game()
    if game:
        print(f"경기 정보: {game}")
        upload_to_sheets(game)
    else:
        print("적재할 경기 데이터 없음")


if __name__ == "__main__":
    main()
