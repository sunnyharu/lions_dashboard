"""
삼성 라이온즈 경기 결과 백필 스크립트
지정 기간의 전체 경기를 KBO API에서 조회해 Google Sheets '경기현황' 탭에 적재
"""
import json
import os
import re
from datetime import datetime, date

import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
import gspread

load_dotenv()

# ── 설정 ────────────────────────────────────────────────
SPREADSHEET_ID    = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"
SHEET_NAME        = "경기현황"
GOOGLE_CREDS_ENV  = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE = "google_credentials.json"

KBO_API = "https://www.koreabaseball.com/ws/Schedule.asmx/GetScheduleList"

# ── 백필 기간 (환경변수 우선, 없으면 기본값) ─────────────────
_start_env = os.environ.get("BACKFILL_START", "2026-04-01")
_end_env   = os.environ.get("BACKFILL_END",   "2026-04-27")
BACKFILL_START = date.fromisoformat(_start_env)
BACKFILL_END   = date.fromisoformat(_end_env)


def fetch_month_games(year: int, month: int) -> list:
    """특정 월의 삼성 라이온즈 경기 전체 조회"""
    start_day = f"{year}{month:02d}01"
    # 해당 월의 마지막 날
    if month == 12:
        end_day = f"{year}1231"
    else:
        import calendar
        last = calendar.monthrange(year, month)[1]
        end_day = f"{year}{month:02d}{last:02d}"

    payload = {
        "leId":      "1",
        "srIdList":  "0,9",
        "seasonId":  str(year),
        "startDay":  start_day,
        "endDay":    end_day,
        "gameMonth": f"{month:02d}",
        "teamId":    "SS",
    }
    resp = requests.post(KBO_API, data=payload, timeout=15)
    resp.raise_for_status()
    data = resp.json()

    games = []
    for entry in data.get("rows", []):
        cells = entry.get("row", [])
        day_cell  = next((c for c in cells if c.get("Class") == "day"),  None)
        play_cell = next((c for c in cells if c.get("Class") == "play"), None)
        if not day_cell or not play_cell:
            continue

        # 날짜 파싱: "04.01(수)" → date(2026, 4, 1)
        day_raw = re.sub(r"\(.*?\)", "", day_cell.get("Text", "")).strip()  # "04.01"
        try:
            mm, dd = day_raw.split(".")
            game_date = date(year, int(mm), int(dd))
        except Exception:
            continue

        if not (BACKFILL_START <= game_date <= BACKFILL_END):
            continue

        parsed = parse_play(play_cell["Text"], game_date)
        if parsed:
            games.append(parsed)

    return games


def parse_play(html: str, game_date: date):
    soup = BeautifulSoup(html, "html.parser")
    spans = soup.find_all("span", recursive=False)
    if len(spans) < 2:
        return None

    away_team = spans[0].get_text(strip=True)
    home_team = spans[-1].get_text(strip=True)

    score_elems = [
        s for s in soup.select("em span")
        if re.fullmatch(r"\d+", s.get_text(strip=True))
    ]
    away_score_cls = score_elems[0].get("class", [""])[0] if len(score_elems) >= 1 else ""
    home_score_cls = score_elems[1].get("class", [""])[0] if len(score_elems) >= 2 else ""

    # 경기 취소/우천: score_elems 없거나 class 없음
    if not away_score_cls and not home_score_cls:
        result    = "취소"
        home_away = "홈" if home_team == "삼성" else "어웨이"
        opponent  = away_team if home_team == "삼성" else home_team
    elif home_team == "삼성":
        home_away = "홈"
        opponent  = away_team
        result    = {"win": "승", "lose": "패", "draw": "무"}.get(home_score_cls, "-")
    else:
        home_away = "어웨이"
        opponent  = home_team
        result    = {"win": "승", "lose": "패", "draw": "무"}.get(away_score_cls, "-")

    return {
        "날짜":     game_date.strftime("%Y.%m.%d"),
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


def upload_all(games: list):
    client = get_gspread_client()
    sh     = client.open_by_key(SPREADSHEET_ID)

    try:
        ws = sh.worksheet(SHEET_NAME)
        ws.clear()
        print(f"기존 '{SHEET_NAME}' 시트 초기화")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=500, cols=10)
        print(f"'{SHEET_NAME}' 시트 신규 생성")

    ws.append_row(["날짜", "홈/어웨이", "상대팀", "결과"])

    rows = [[g["날짜"], g["홈/어웨이"], g["상대팀"], g["결과"]] for g in games]
    if rows:
        ws.append_rows(rows)

    print(f"적재 완료: {len(rows)}경기")
    for r in rows:
        print(f"  {r[0]} | {r[1]:5} | {r[2]:4} | {r[3]}")


def main():
    print(f"백필 기간: {BACKFILL_START} ~ {BACKFILL_END}")
    # 4월만 조회 (기간이 한 달 내)
    games = fetch_month_games(BACKFILL_START.year, BACKFILL_START.month)
    games.sort(key=lambda g: g["날짜"])
    print(f"조회된 경기 수: {len(games)}")

    if games:
        upload_all(games)
    else:
        print("조회된 경기 없음")


if __name__ == "__main__":
    main()
