"""
특이사항 업데이트: 일별매출 시트의 특이사항 컬럼을 업데이트

환경변수:
  INPUT_DATE          : 날짜 (YYYY.MM.DD)
  INPUT_NOTE          : 특이사항 내용
  GOOGLE_CREDENTIALS  : Google 서비스 계정 JSON
"""
import json
import os
from google.oauth2.service_account import Credentials
import gspread
from dotenv import load_dotenv

load_dotenv()

INPUT_DATE        = os.environ.get("INPUT_DATE", "").strip()
INPUT_NOTE        = os.environ.get("INPUT_NOTE", "").strip()
GOOGLE_CREDS_ENV  = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE = "google_credentials.json"
SPREADSHEET_ID    = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"
SHEET_NAME        = "일별매출"


def normalize_date(s: str) -> str:
    try:
        parts = s.split(".")
        return f"{parts[0]}.{int(parts[1]):02d}.{int(parts[2]):02d}"
    except Exception:
        return s


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


def main():
    if not INPUT_DATE:
        print("오류: INPUT_DATE 없음")
        return
    if not INPUT_NOTE:
        print("오류: INPUT_NOTE 없음")
        return

    date_key = normalize_date(INPUT_DATE)
    print(f"특이사항 업데이트: {date_key} → {INPUT_NOTE}")

    client = get_gspread_client()
    sh     = client.open_by_key(SPREADSHEET_ID)
    ws     = sh.worksheet(SHEET_NAME)

    existing = ws.get_all_values()
    if not existing:
        print("시트가 비어 있음")
        return

    header = existing[0]
    date_col_idx = header.index("날짜") if "날짜" in header else 0

    # 특이사항 컬럼 확인 / 없으면 추가
    if "특이사항" not in header:
        header.append("특이사항")
        note_col_idx = len(header) - 1
        ws.update([header], "1:1")
        print("'특이사항' 컬럼 추가")
    else:
        note_col_idx = header.index("특이사항")

    note_col_letter = chr(65 + note_col_idx)

    # 해당 날짜 행 찾기
    for i, row in enumerate(existing[1:], start=2):
        if row and len(row) > date_col_idx:
            if normalize_date(row[date_col_idx]) == date_key:
                ws.update([[INPUT_NOTE]], f"{note_col_letter}{i}")
                print(f"  완료: {note_col_letter}{i} 셀 업데이트")
                return

    # 날짜 행이 없으면 새 행 추가
    new_row = [""] * len(header)
    new_row[date_col_idx]  = date_key
    new_row[note_col_idx]  = INPUT_NOTE
    ws.append_row(new_row, value_input_option="USER_ENTERED")
    print(f"  신규 행 추가: {date_key}")


if __name__ == "__main__":
    main()
