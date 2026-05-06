"""
Gmail에서 Berriz Shop Report PDF를 수신 → 결제금액(ON거래액) 추출 → 일별매출 시트 업데이트

환경변수:
  GOOGLE_EMAIL        : Gmail 계정 (ex: freelywind222@gmail.com)
  GMAIL_APP_PASSWORD  : Gmail 앱 비밀번호 (16자리, 공백 없이)
  GOOGLE_CREDENTIALS  : Google 서비스 계정 JSON (시트 업데이트용)
"""
import email
import email.header
import imaplib
import io
import json
import os
import re
from datetime import datetime, timezone, timedelta

from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
import gspread
import pdfplumber

load_dotenv()

GMAIL_USER        = os.environ.get("GOOGLE_EMAIL", "")
GMAIL_APP_PASS    = os.environ.get("GMAIL_APP_PASSWORD", "")
GOOGLE_CREDS_ENV  = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE = "google_credentials.json"
SPREADSHEET_ID    = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"
SHEET_NAME        = "일별매출"

KST = timezone(timedelta(hours=9))


# ── 유틸 ──────────────────────────────────────────────────────────────────────

def normalize_date(s: str) -> str:
    """2026.5.3 → 2026.05.03"""
    try:
        parts = s.split(".")
        return f"{parts[0]}.{int(parts[1]):02d}.{int(parts[2]):02d}"
    except Exception:
        return s


def decode_header_value(raw) -> str:
    parts = email.header.decode_header(raw or "")
    result = ""
    for part, enc in parts:
        if isinstance(part, bytes):
            result += part.decode(enc or "utf-8", errors="ignore")
        else:
            result += str(part)
    return result


# ── Google Sheets ──────────────────────────────────────────────────────────────

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


# ── PDF 파싱 ───────────────────────────────────────────────────────────────────

def parse_pdf(pdf_bytes: bytes) -> list:
    """
    PDF에서 날짜·결제금액 추출
    반환: [{"date": "2026.05.05", "amount": 45165500}, ...]
    """
    results = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        full_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"

    print("=== PDF 텍스트 (앞 2000자) ===")
    print(full_text[:2000])
    print("==============================\n")

    # 연도 추출: "Apr 29, 2026 - May 5, 2026" 또는 "2026" 포함 문자열
    year_match = re.search(r"\b(20\d{2})\b", full_text)
    year = int(year_match.group(1)) if year_match else datetime.now(KST).year

    # 행 패턴: MM/DD <결제금액> <건당평균> <배송비제외금액>
    # 예: "05/05 45,165,500 96,714.1 44,778,500"
    row_pattern = re.compile(r"(\d{2}/\d{2})\s+([\d,]+)\s+[\d,.]+\s+[\d,]+")

    for m in row_pattern.finditer(full_text):
        month_day  = m.group(1)   # "05/05"
        amount_str = m.group(2)   # "45,165,500"
        month, day = map(int, month_day.split("/"))
        amount = int(amount_str.replace(",", ""))
        date_str = f"{year}.{month:02d}.{day:02d}"
        results.append({"date": date_str, "amount": amount})
        print(f"  파싱: {date_str} → {amount:,}원")

    return results


# ── Gmail IMAP ────────────────────────────────────────────────────────────────

def fetch_latest_pdf_from_gmail() -> bytes:
    """
    Gmail IMAP으로 최근 Berriz Shop Report PDF 첨부파일 수신
    검색 우선순위:
      1. 제목/파일명에 LIONS 또는 BERRIZ 포함
      2. 루커 스튜디오 발신 (google/looker 도메인)
    """
    if not GMAIL_USER or not GMAIL_APP_PASS:
        print("오류: GOOGLE_EMAIL 또는 GMAIL_APP_PASSWORD 환경변수 없음")
        return None

    print(f"Gmail IMAP 접속 중: {GMAIL_USER}")
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        mail.login(GMAIL_USER, GMAIL_APP_PASS)
        mail.select("INBOX")
        print("  로그인 성공")
    except Exception as e:
        print(f"  Gmail 로그인 실패: {e}")
        return None

    # 키워드 검색
    search_terms = [
        'SUBJECT "LIONS"',
        'SUBJECT "Berriz"',
        'SUBJECT "베리즈"',
        'SUBJECT "Shop Report"',
    ]
    uids = []
    for term in search_terms:
        try:
            _, data = mail.search(None, term)
            found = data[0].split()
            if found:
                uids.extend(found)
                print(f"  검색 '{term}': {len(found)}건")
        except Exception:
            pass

    # 키워드 검색 결과 없으면 최근 30통에서 PDF 탐색
    if not uids:
        print("  키워드 검색 없음 → 최근 30통 탐색")
        _, data = mail.search(None, "ALL")
        all_uids = data[0].split()
        uids = all_uids[-30:]

    # 중복 제거, 최신순
    seen = set()
    unique_uids = []
    for u in reversed(uids):
        if u not in seen:
            seen.add(u)
            unique_uids.append(u)

    pdf_data = None
    best_pdf  = None  # 키워드 매칭된 PDF 우선

    for uid in unique_uids:
        try:
            _, msg_data = mail.fetch(uid, "(RFC822)")
            msg = email.message_from_bytes(msg_data[0][1])
            subject = decode_header_value(msg.get("Subject", ""))
            sender  = msg.get("From", "")

            for part in msg.walk():
                ct = part.get_content_type()
                raw_filename = part.get_filename() or ""
                filename = decode_header_value(raw_filename)

                is_pdf = ct == "application/pdf" or filename.lower().endswith(".pdf")
                if not is_pdf:
                    continue

                combined = (filename + " " + subject).upper()
                is_match = any(k in combined for k in ["LIONS", "BERRIZ", "SHOP_REPORT", "SHOP REPORT"])

                payload = part.get_payload(decode=True)
                if is_match:
                    print(f"  ✓ PDF 매칭: '{filename}' | 제목: {subject[:50]} | 발신: {sender[:40]}")
                    best_pdf = payload
                    break
                else:
                    if pdf_data is None:
                        print(f"  PDF 후보: '{filename}' | 발신: {sender[:40]}")
                        pdf_data = payload

            if best_pdf:
                break  # 최적 매칭 발견

        except Exception as e:
            print(f"  메일 처리 오류: {e}")
            continue

    mail.logout()

    result = best_pdf or pdf_data
    if result:
        print(f"  PDF 수신 완료: {len(result):,} bytes")
    else:
        print("  Berriz Shop Report PDF를 찾지 못했습니다.")
    return result


# ── Google Sheets 업데이트 ─────────────────────────────────────────────────────

def update_sheets(rows: list):
    """
    rows: [{"date": "2026.05.05", "amount": 45165500}, ...]
    일별매출 시트의 ON거래액 컬럼 업데이트 (없으면 행 추가)
    """
    if not rows:
        print("업데이트할 데이터 없음")
        return

    client = get_gspread_client()
    sh     = client.open_by_key(SPREADSHEET_ID)
    ws     = sh.worksheet(SHEET_NAME)

    existing = ws.get_all_values()
    if not existing:
        print("시트가 비어 있음")
        return

    header = list(existing[0])

    # ON거래액 컬럼 확인 / 없으면 추가
    if "ON거래액" not in header:
        header.append("ON거래액")
        on_col_idx = len(header) - 1
        ws.update([header], "1:1")
        print(f"'ON거래액' 컬럼 추가 (열 {on_col_idx + 1})")
    else:
        on_col_idx = header.index("ON거래액")

    date_col_idx  = header.index("날짜") if "날짜" in header else 0
    on_col_letter = chr(65 + on_col_idx)

    # 기존 날짜 → (행번호, 행데이터) 매핑
    existing_map = {}
    for i, r in enumerate(existing[1:], start=2):
        if r and len(r) > date_col_idx and r[date_col_idx].strip():
            existing_map[normalize_date(r[date_col_idx])] = (i, r)

    inserted = updated = skipped = 0
    for row in rows:
        date_key = normalize_date(row["date"])
        amount   = row["amount"]

        if date_key in existing_map:
            row_num, existing_row = existing_map[date_key]
            existing_on = existing_row[on_col_idx].strip() if len(existing_row) > on_col_idx else ""
            if not existing_on or existing_on in ("0", ""):
                ws.update([[amount]], f"{on_col_letter}{row_num}")
                print(f"  ON거래액 업데이트: {date_key} → {amount:,}원")
                updated += 1
            else:
                print(f"  ON거래액 이미 존재 스킵: {date_key} ({existing_on})")
                skipped += 1
        else:
            # 새 행: 날짜 + ON거래액
            new_row = [""] * len(header)
            new_row[date_col_idx] = date_key
            new_row[on_col_idx]   = amount
            ws.append_row(new_row, value_input_option="USER_ENTERED")
            existing_map[date_key] = (0, new_row)
            print(f"  신규 행 추가: {date_key} → {amount:,}원")
            inserted += 1

    print(f"\n완료: {inserted}행 추가 / {updated}행 업데이트 / {skipped}행 스킵")


# ── 메인 ──────────────────────────────────────────────────────────────────────

def main():
    print("=== Berriz Shop Report → ON거래액 업데이트 ===\n")

    # 1. Gmail에서 PDF 수신
    pdf_bytes = fetch_latest_pdf_from_gmail()
    if not pdf_bytes:
        print("PDF 없음. 종료.")
        return

    # 2. PDF 파싱
    print("\n[PDF 파싱]")
    rows = parse_pdf(pdf_bytes)
    if not rows:
        print("데이터 추출 실패. PDF 텍스트 구조를 확인하세요.")
        return
    print(f"추출된 날짜·금액: {len(rows)}건")

    # 3. Google Sheets 업데이트
    print("\n[Sheets 업데이트]")
    update_sheets(rows)


if __name__ == "__main__":
    main()
