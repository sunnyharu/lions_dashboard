"""
어제 날짜 데이터가 이미 시트에 있으면 GITHUB_OUTPUT에 skip=true 출력
→ GitHub Actions에서 이후 스텝을 건너뜀
"""
import json
import os
import sys
from datetime import datetime, timedelta

from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
import gspread

load_dotenv()

SPREADSHEET_ID   = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"
GOOGLE_CREDS_ENV = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE = "google_credentials.json"

yesterday = datetime.today() - timedelta(days=1)
YESTERDAY_STR = yesterday.strftime("%Y.%m.%d")  # "2026.04.28"


def set_output(key: str, value: str):
    github_output = os.environ.get("GITHUB_OUTPUT", "")
    if github_output:
        with open(github_output, "a") as f:
            f.write(f"{key}={value}\n")
    else:
        print(f"::set-output name={key}::{value}")


def main():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    try:
        if GOOGLE_CREDS_ENV:
            creds = Credentials.from_service_account_info(json.loads(GOOGLE_CREDS_ENV), scopes=scopes)
        else:
            creds = Credentials.from_service_account_file(GOOGLE_CREDS_FILE, scopes=scopes)
        client = gspread.authorize(creds)
        sh     = client.open_by_key(SPREADSHEET_ID)
        ws     = sh.worksheet("일별매출")
        vals   = ws.get_all_values()
    except Exception as e:
        print(f"시트 조회 오류: {e} → 스킵 없이 실행")
        set_output("skip", "false")
        return

    header = vals[0] if vals else []
    date_col = header.index("날짜") if "날짜" in header else 0

    existing_dates = {
        row[date_col].strip()
        for row in vals[1:]
        if row and len(row) > date_col
    }

    if YESTERDAY_STR in existing_dates:
        print(f"이미 실행됨: {YESTERDAY_STR} 데이터 존재 → 스킵")
        set_output("skip", "true")
    else:
        print(f"미실행: {YESTERDAY_STR} 데이터 없음 → 실행")
        set_output("skip", "false")


if __name__ == "__main__":
    main()
