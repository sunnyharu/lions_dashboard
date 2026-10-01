"""
삼성 라이온즈 공식 홈페이지 선수단 → Google Sheets '선수명단' 탭 (주 1회)

대시보드 상품 실적 분석의 '선수별'은 이 명단에 있는 이름만 선수로 인정한다.
(상품 옵션에는 '타자'·'투수'·마스코트명처럼 선수가 아닌 값도 섞여 있다)

- 대상: 투수 / 타자 / 신입단 / 군입대 / 재활 (코칭스태프 제외)
- 명단은 덮어쓰지 않고 누적한다. 이적·방출 선수의 과거 굿즈 실적도 선수로 남겨야 해서
  최초확인·최근확인 일자만 갱신한다.
"""
import json
import os
import re
import urllib.request
from datetime import datetime, timedelta, timezone

from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
import gspread

load_dotenv()

SPREADSHEET_ID    = os.environ.get("SPREADSHEET_ID", "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU")
GOOGLE_CREDS_ENV  = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE = "google_credentials.json"
SHEET_NAME        = "선수명단"
HEADER            = ["이름", "등번호", "포지션", "구분", "최초확인", "최근확인"]

PAGES = {2: "투수", 3: "타자", 5: "신입단", 4: "군입대", 6: "재활"}
URL   = "https://www.samsunglions.com/roster/roster_{}_list.asp"
# 공식 선수단 페이지에 없지만 굿즈가 팔리는 선수 - 은퇴 레전드와 시즌 중 떠난 선수.
# 굿즈 실적이 생기면 여기에 이름을 더한다.
EXTRA = {
    "레전드":  ["오승환", "이승엽", "양준혁", "이만수"],
    "前 선수": ["매닝", "미야지"],
}
TODAY = datetime.now(timezone(timedelta(hours=9))).strftime("%Y.%m.%d")


def fetch(page: int) -> str:
    req = urllib.request.Request(URL.format(page), headers={"User-Agent": "Mozilla/5.0"})
    return urllib.request.urlopen(req, timeout=30).read().decode("utf-8", "replace")


def parse(html: str, group: str) -> list:
    """[(이름, 등번호, 포지션)] - 포지션은 소제목(내야수 등), 없으면 구분명."""
    players = []
    # 소제목(<h4><em>내야수</em> 총 …) 단위로 잘라 포지션을 붙인다
    parts = re.split(r'<h4>\s*<em>([^<]+)</em>', html)
    sections = [(group, parts[0])] + [(parts[i].strip(), parts[i + 1]) for i in range(1, len(parts) - 1, 2)]
    for pos, chunk in sections:
        for raw in re.findall(r'class="na">\s*([^<]+?)\s*</', chunk):
            m = re.match(r"(\d{1,3})\.\s*(.+)", raw)
            no, name = (m.group(1), m.group(2)) if m else ("", raw)
            name = name.strip()
            if re.fullmatch(r"[가-힣A-Za-z]{2,10}", name):
                players.append((name, no, pos))
    if not players:   # 군입대 페이지처럼 이름 옆에 '전역예정'만 붙는 형식
        text = re.sub(r"<[^>]+>", " ", html)
        for name in re.findall(r"([가-힣]{2,5})\s+\d{4}-\d{2}-\d{2}\s*전역예정", text):
            players.append((name, "", group))
    return players


def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    if GOOGLE_CREDS_ENV:
        creds = Credentials.from_service_account_info(json.loads(GOOGLE_CREDS_ENV), scopes=scopes)
    else:
        creds = Credentials.from_service_account_file(GOOGLE_CREDS_FILE, scopes=scopes)
    return gspread.authorize(creds)


def main():
    found = {}
    for page, group in PAGES.items():
        got = parse(fetch(page), group)
        print(f"  {group}: {len(got)}명")
        for name, no, pos in got:
            found.setdefault((name, no), (name, no, pos, group))
    for group, names in EXTRA.items():
        for name in names:
            if not any(k[0] == name for k in found):
                found[(name, "")] = (name, "", group, group)
    if len(found) < 30:   # 페이지 구조가 바뀌어 거의 못 읽었으면 명단을 망가뜨리지 않고 실패로 끝낸다
        raise SystemExit(f"수집 {len(found)}명 - 페이지 구조 변경 의심, 시트를 갱신하지 않음")

    sh = get_client().open_by_key(SPREADSHEET_ID)
    try:
        ws = sh.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=500, cols=len(HEADER))
    rows = ws.get_all_values()
    existing = {(r[0], r[1]): r for r in rows[1:] if r and r[0]}

    out, new = [], 0
    for key, r in existing.items():
        if key in found:
            name, no, pos, group = found.pop(key)
            out.append([name, no, pos, group, r[4] or TODAY, TODAY])
        else:
            out.append(r[:6] + [""] * (6 - len(r[:6])))   # 이번엔 없지만 과거 선수로 유지
    for name, no, pos, group in found.values():
        out.append([name, no, pos, group, TODAY, TODAY]); new += 1
    out.sort(key=lambda r: (r[3], r[2], r[0]))
    ws.clear()
    ws.update([HEADER] + out, "A1", raw=True)
    active = sum(1 for r in out if r[5] == TODAY)
    print(f"선수명단 갱신: 전체 {len(out)}명 (이번 확인 {active}명, 신규 {new}명)")


if __name__ == "__main__":
    main()
