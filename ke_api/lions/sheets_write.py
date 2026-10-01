"""대시보드에서 들어오는 쓰기 요청 — 특이사항(일별매출)·온라인 상품별 업로드(상품별매출(on)).

예전에는 특이사항이 Cloudflare Worker → GitHub Actions → update_note.py 로,
업로드가 Google Apps Script 웹앱으로 처리됐다. 같은 시트 동작을 API 안에서 직접 수행한다.
"""
from . import generate_dashboard as gd

NOTE_SHEET = "일별매출"
UPLOAD_SHEET = "상품별매출(on)"
UPLOAD_HEADER = ["판매일자", "상품ID", "상품명", "바코드", "skucode", "사이즈", "선수명", "판매단가", "판매수량", "실판매금액"]


def _norm_date(s: str) -> str:
    try:
        p = s.strip().replace("-", ".").split(".")
        return f"{p[0]}.{int(p[1]):02d}.{int(p[2]):02d}"
    except Exception:
        return s.strip()


def _col(idx: int) -> str:
    out = ""
    idx += 1
    while idx:
        idx, r = divmod(idx - 1, 26)
        out = chr(65 + r) + out
    return out


def write_note(date: str, note: str) -> str:
    """일별매출의 해당 날짜 특이사항 셀을 갱신, 날짜 행이 없으면 새 행을 추가한다."""
    ws = gd.get_client().open_by_key(gd.SPREADSHEET_ID).worksheet(NOTE_SHEET)
    rows = ws.get_all_values()
    header = rows[0]
    di = header.index("날짜") if "날짜" in header else 0
    if "특이사항" not in header:
        header.append("특이사항")
        ws.update([header], "1:1")
    ni = header.index("특이사항")
    key = _norm_date(date)
    for i, r in enumerate(rows[1:], start=2):
        if r and len(r) > di and _norm_date(r[di]) == key:
            ws.update([[note]], f"{_col(ni)}{i}")
            return "updated"
    new = [""] * len(header)
    new[di], new[ni] = key, note
    ws.append_row(new, value_input_option="USER_ENTERED")
    return "appended"


def upload_online_rows(rows: list) -> dict:
    """업로드된 날짜의 기존 행을 지우고 새 행으로 교체한다 (Apps Script doPost 와 동일 규칙)."""
    upload_dates = {str(r.get("판매일자", "")).strip() for r in rows}
    ws = gd.get_client().open_by_key(gd.SPREADSHEET_ID).worksheet(UPLOAD_SHEET)
    allv = ws.get_all_values()
    has_header = bool(allv) and allv[0] and allv[0][0] == UPLOAD_HEADER[0]
    existing = allv[1:] if has_header else allv
    kept = [r for r in existing if str(r[0] if r else "").strip().replace(".", "-") not in upload_dates]
    new = [[str(r.get(c, "") if r.get(c) is not None else "") for c in UPLOAD_HEADER] for r in rows]
    data = [UPLOAD_HEADER] + kept + new
    # 먼저 비우고 나눠 쓰면 중간 실패 시 시트가 빈 채로 남는다 → 덮어쓴 뒤 남는 꼬리만 지운다
    if len(data) > ws.row_count:
        ws.add_rows(len(data) - ws.row_count)
    for i in range(0, len(data), 10000):
        ws.update(data[i:i + 10000], f"A{i + 1}", raw=True)
    if len(allv) > len(data):
        ws.batch_clear([f"A{len(data) + 1}:{_col(len(UPLOAD_HEADER) - 1)}{len(allv)}"])
    return {"inserted": len(new), "dates": sorted(upload_dates)}
