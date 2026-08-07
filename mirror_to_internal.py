"""
개인 시트(playmd) → 사내 시트(playmd_사내이관테스트) 일일 미러링
- 원본은 읽기만, 사내 시트만 덮어씀
- 사내 이관 병행 운영용: 사내 시트가 매일 최신으로 유지되는지 검증
"""
import json
import os
import time

from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials

load_dotenv()

SRC_ID = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"   # 개인(운영)
DST_ID = "1kwsFg2DoIlyYr5tfpx0b1htq0PSXpV1gqlls496Pb6s"   # 사내(테스트)

GOOGLE_CREDS_ENV  = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE = "google_credentials.json"

# 미러링 제외 탭 (2026-08-07: 개인 시트도 순결제 기준으로 전환돼 제외 없음)
EXCLUDE = set()


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


def main():
    client = get_client()
    src = client.open_by_key(SRC_ID)
    dst = client.open_by_key(DST_ID)
    dst_tabs = {ws.title: ws for ws in dst.worksheets()}

    for ws in src.worksheets():
        name = ws.title
        if name in EXCLUDE:
            print(f"[{name}] 제외 (사내 기준 상이)")
            continue

        vals = ws.get_all_values()
        rows = len(vals)
        cols = max((len(r) for r in vals), default=1)

        if name in dst_tabs:
            target = dst_tabs[name]
            target.clear()
            target.resize(rows=max(rows, 1), cols=max(cols, 1))
        else:
            target = dst.add_worksheet(title=name, rows=max(rows, 1), cols=max(cols, 1))

        if vals:
            for i in range(0, rows, 10000):
                target.update(vals[i:i + 10000], f"A{i + 1}", raw=True)
                time.sleep(1)
        print(f"[{name}] 미러링 완료 ({rows}행)")
        time.sleep(1)

    print("사내 시트 미러링 완료")


if __name__ == "__main__":
    main()
