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

KBO_API = "https://www.koreabaseball.com/ws/Schedule.asmx/GetScheduleList"


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
    resp = requests.post(KBO_API, data=payload, timeout=15)
    resp.raise_for_status()
    data = resp.json()

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

        return parse_play(play_cell["Text"])

    print(f"어제({DAY_TEXT}) 삼성 라이온즈 경기 없음")
    return None


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
        ws.append_row(["날짜", "홈/어웨이", "상대팀", "결과"])

    row = [game["날짜"], game["홈/어웨이"], game["상대팀"], game["결과"]]
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
