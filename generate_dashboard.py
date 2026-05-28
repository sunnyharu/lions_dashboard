"""
Google Sheets 두 탭(일별매출 + 경기현황)을 읽어 dashboard/index.html 생성
"""
import json
import os
from datetime import datetime

from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
import gspread

load_dotenv()

SPREADSHEET_ID    = "1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU"
GOOGLE_CREDS_ENV  = os.environ.get("GOOGLE_CREDENTIALS", "")
GOOGLE_CREDS_FILE = "google_credentials.json"


def get_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    if GOOGLE_CREDS_ENV:
        info  = json.loads(GOOGLE_CREDS_ENV)
        creds = Credentials.from_service_account_info(info, scopes=scopes)
    else:
        creds = Credentials.from_service_account_file(GOOGLE_CREDS_FILE, scopes=scopes)
    return gspread.authorize(creds)


def fetch_data():
    client = get_client()
    sh     = client.open_by_key(SPREADSHEET_ID)

    # ── 일별매출 ──────────────────────────────────────────
    ws_sales  = sh.worksheet("일별매출")
    sales_raw = ws_sales.get_all_records()
    sales = {}
    for row in sales_raw:
        date = str(row.get("날짜", "")).strip()
        if not date:
            continue
        # 날짜 정규화
        parts = date.split(".")
        if len(parts) == 3:
            date = f"{parts[0]}.{int(parts[1]):02d}.{int(parts[2]):02d}"
        off  = row.get("OFF거래액", "") or 0
        on   = row.get("ON거래액",  "") or 0
        note = str(row.get("특이사항", "") or "").strip()
        try: off = int(float(str(off).replace(",", "")))
        except: off = 0
        try: on  = int(float(str(on).replace(",", "")))
        except: on  = 0
        # 날짜 중복 시 ON거래액이 있는 행 우선
        if date in sales and on == 0:
            continue
        sales[date] = {"off": off, "on": on, "note": note}

    # ── 뉴스이슈 ──────────────────────────────────────────
    news = []
    try:
        ws_news   = sh.worksheet("뉴스이슈")
        all_vals  = ws_news.get_all_values()
        if len(all_vals) > 1:
            header = all_vals[0]
            def col(h, fallback=-1):
                return header.index(h) if h in header else fallback

            ci_date    = col("날짜",   0)
            ci_source  = col("출처",   1)
            ci_title   = col("제목",   2)
            ci_summary = col("AI요약", col("요약", 3))
            ci_views   = col("조회수", -1)
            ci_link    = col("링크",   col("링크", 4))

            for row in all_vals[1:]:
                def cell(i, default=""):
                    return row[i].strip() if i >= 0 and i < len(row) else default

                date = cell(ci_date)
                if not date:
                    continue
                parts = date.split(".")
                if len(parts) == 3:
                    try:
                        date = f"{parts[0]}.{int(parts[1]):02d}.{int(parts[2]):02d}"
                    except Exception:
                        pass
                views = 0
                if ci_views >= 0:
                    try: views = int(cell(ci_views, "0") or 0)
                    except: views = 0
                news.append({
                    "date":    date,
                    "source":  cell(ci_source),
                    "title":   cell(ci_title),
                    "summary": cell(ci_summary),
                    "views":   views,
                    "link":    cell(ci_link),
                })
        print(f"뉴스이슈 {len(news)}건 로드")
    except Exception as e:
        print(f"뉴스이슈 로드 오류: {e}")

    # ── 카페트렌드 다이제스트 ─────────────────────────────
    digest = ""
    try:
        ws_digest  = sh.worksheet("카페트렌드")
        digest_raw = ws_digest.get_all_values()
        if len(digest_raw) > 1:
            digest = digest_raw[-1][1] if len(digest_raw[-1]) > 1 else ""
    except Exception:
        pass

    # ── 상품별매출 (row별 수집) ──────────────────────────────
    raw_products = []
    try:
        ws_prod   = sh.worksheet("상품별매출")
        prod_raw  = ws_prod.get_all_records()
        for row in prod_raw:
            date   = str(row.get("판매일자", "") or "").strip()
            code   = str(row.get("상품코드", "") or "").strip()
            name   = str(row.get("상품명",   "") or "").strip()
            color  = str(row.get("칼라명",   "") or "").strip()
            size   = str(row.get("사이즈명", "") or "").strip()
            if not code and not name:
                continue
            def _int(v):
                try: return int(float(str(v).replace(",", "")))
                except: return 0
            price  = _int(row.get("판매단가",  0))
            qty    = _int(row.get("판매수량",  0))
            amount = _int(row.get("실판매금액", 0))
            raw_products.append({
                "date":   date,
                "code":   code,
                "name":   name,
                "color":  color,
                "size":   size,
                "price":  price,
                "qty":    qty,
                "amount": amount,
            })
        print(f"상품별매출 {len(raw_products)}행 로드")
    except Exception as e:
        print(f"상품별매출 로드 오류: {e}")

    # ── 경기현황 ──────────────────────────────────────────
    ws_game  = sh.worksheet("경기현황")
    game_raw = ws_game.get_all_records()
    games = {}
    for row in game_raw:
        date = str(row.get("날짜", "")).strip()
        if not date:
            continue
        parts = date.split(".")
        if len(parts) == 3:
            date = f"{parts[0]}.{int(parts[1]):02d}.{int(parts[2]):02d}"
        crowd = 0
        try: crowd = int(float(str(row.get("관중수", 0) or 0).replace(",", "")))
        except: crowd = 0
        games[date] = {
            "home_away": row.get("홈/어웨이", ""),
            "opponent":  row.get("상대팀", ""),
            "result":    row.get("결과", ""),
            "crowd":     crowd,
        }

    # ── 병합 ─────────────────────────────────────────────
    all_dates = sorted(set(list(sales.keys()) + list(games.keys())))
    merged = []
    for d in all_dates:
        s = sales.get(d, {"off": 0, "on": 0})
        g = games.get(d, {"home_away": "", "opponent": "", "result": ""})
        merged.append({
            "date":      d,
            "off":       s["off"],
            "on":        s["on"],
            "total":     s["off"] + s["on"],
            "home_away": g["home_away"],
            "opponent":  g["opponent"],
            "result":    g["result"],
            "crowd":     g.get("crowd", 0),
            "note":      s.get("note", ""),
        })

    return merged, news, digest, raw_products


def build_html(data: list, news: list, digest: str, raw_products: list) -> str:
    game_days = [r for r in data if r["result"] and r["result"] != "취소"]

    def avg(lst): return int(sum(lst) / len(lst)) if lst else 0
    def fmt(n):   return f"{n:,}" if n else "-"

    total_off   = sum(r["off"]   for r in data if r["off"])
    total_on    = sum(r["on"]    for r in data if r["on"])
    total_total = total_off + total_on

    dates_with_sales = [r["date"] for r in data if r["off"] or r["on"]]
    date_range_start = dates_with_sales[0]  if dates_with_sales else ""
    date_range_end   = dates_with_sales[-1] if dates_with_sales else ""
    date_range_label = f"{date_range_start} ~ {date_range_end}" if date_range_start else ""

    home_rows  = [r for r in game_days if r["home_away"] == "홈"]
    away_rows  = [r for r in game_days if r["home_away"] == "어웨이"]
    win_rows   = [r for r in game_days if r["result"] == "승"]
    lose_rows  = [r for r in game_days if r["result"] == "패"]
    wins       = len(win_rows)
    losses     = len(lose_rows)
    win_rate   = f"{wins/(wins+losses)*100:.0f}%" if (wins+losses) > 0 else "-"

    # 관중수 통계 (홈 경기만)
    STADIUM_CAPACITY = 24000
    home_crowd_rows  = [r for r in home_rows if r.get("crowd", 0) > 0]
    avg_crowd        = int(sum(r["crowd"] for r in home_crowd_rows) / len(home_crowd_rows)) if home_crowd_rows else 0
    avg_occupancy    = f"{avg_crowd / STADIUM_CAPACITY * 100:.1f}%" if avg_crowd else "-"
    total_crowd      = sum(r["crowd"] for r in home_crowd_rows)

    # Chart 데이터
    chart_dates  = [r["date"][5:] for r in data]  # "04.01" 형태
    chart_off    = [r["off"]   for r in data]
    chart_on     = [r["on"]    for r in data]
    chart_total  = [r["total"] for r in data]
    chart_notes  = [r.get("note", "") for r in data]

    # 월별 홈/어웨이 평균
    from collections import defaultdict
    def month_label(d): return d[:4] + "." + d[5:7]  # "2026.04"

    monthly_ha  = defaultdict(lambda: defaultdict(list))
    monthly_res = defaultdict(lambda: defaultdict(list))
    for r in game_days:
        m = month_label(r["date"])
        ha = r["home_away"]
        rs = r["result"]
        if ha:
            if r["off"]: monthly_ha[m][ha + "_off"].append(r["off"])
            if r["on"]:  monthly_ha[m][ha + "_on"].append(r["on"])
        if rs in ["승", "패"]:
            if r["off"]: monthly_res[m][rs + "_off"].append(r["off"])
            if r["on"]:  monthly_res[m][rs + "_on"].append(r["on"])

    ha_months  = sorted(monthly_ha.keys())
    res_months = sorted(monthly_res.keys())

    ha_labels, ha_off_data, ha_on_data = [], [], []
    for ha in ["홈", "어웨이"]:
        for m in ha_months:
            ml = m[5:].lstrip("0") + "월"
            ha_labels.append(f"{ha} {ml}")
            ha_off_data.append(avg(monthly_ha[m].get(ha + "_off", [])))
            ha_on_data.append(avg(monthly_ha[m].get(ha + "_on",  [])))

    res_labels, res_off_data, res_on_data = [], [], []
    for rs in ["승", "패"]:
        for m in res_months:
            ml = m[5:].lstrip("0") + "월"
            res_labels.append(f"{rs} {ml}")
            res_off_data.append(avg(monthly_res[m].get(rs + "_off", [])))
            res_on_data.append(avg(monthly_res[m].get(rs + "_on",  [])))

    # ── 월별 요약 데이터 계산 ────────────────────────────────
    monthly_summary = defaultdict(lambda: {"off": 0, "on": 0})
    for r in data:
        m = r["date"][:7]  # "2026.05"
        monthly_summary[m]["off"] += r["off"]
        monthly_summary[m]["on"]  += r["on"]
    unique_months = sorted(monthly_summary.keys())

    # 월별 필터 버튼 HTML
    filter_btns_html = '<button class="filter-btn active" data-month="all">전체</button>\n'
    for m in unique_months:
        label = f"{m[2:4]}년 {int(m[5:7])}월"
        filter_btns_html += f'<button class="filter-btn" data-month="{m}">{label}</button>\n'

    # 월별 요약 테이블 행 HTML
    summary_rows_html = ""
    grand_off = grand_on = 0
    for m in unique_months:
        off = monthly_summary[m]["off"]
        on  = monthly_summary[m]["on"]
        tot = off + on
        grand_off += off
        grand_on  += on
        label = f"{m[2:4]}년 {int(m[5:7])}월"
        summary_rows_html += f"""<tr>
      <td>{label}</td>
      <td class="num">{fmt(off)}</td>
      <td class="num">{fmt(on)}</td>
      <td class="num bold">{fmt(tot)}</td>
    </tr>"""
    grand_tot = grand_off + grand_on
    summary_rows_html += f"""<tr class="summary-total">
  <td>합계</td>
  <td class="num">{fmt(grand_off)}</td>
  <td class="num">{fmt(grand_on)}</td>
  <td class="num bold">{fmt(grand_tot)}</td>
</tr>"""

    raw_products_json = json.dumps(raw_products, ensure_ascii=False)

    # 뉴스 / 카페 분리
    news_items = sorted(
        [n for n in news if n["source"] == "뉴스"],
        key=lambda x: x.get("date", ""), reverse=True
    )[:10]
    from datetime import date as _date, timedelta as _td
    _cutoff = (_date.today() - _td(days=3)).strftime("%Y.%m.%d")
    cafe_items = sorted(
        [n for n in news if "카페" in n["source"] and n.get("date", "") >= _cutoff],
        key=lambda x: x.get("views", 0), reverse=True
    )[:10]

    def issue_item_html(n):
        link_open  = f'<a href="{n["link"]}" target="_blank" rel="noopener">' if n["link"] else ""
        link_close = "</a>" if n["link"] else ""
        views_html = f'<span class="issue-views">조회 {int(n.get("views",0)):,}</span>' if n.get("views") else ""
        return f"""<div class="issue-item">
          <div class="issue-meta">{views_html}</div>
          {link_open}<div class="issue-title">{n["title"]}</div>{link_close}
          {'<div class="issue-desc">' + n["summary"] + '</div>' if n["summary"] else ""}
        </div>"""

    # 왼쪽: 뉴스
    news_col_html = "".join(issue_item_html(n) for n in news_items) or '<div class="issue-empty">수집된 뉴스 없음</div>'

    # 오른쪽: 카페 다이제스트
    digest_col_html = f'<div class="digest-text">{digest}</div>' if digest else '<div class="issue-empty">카페 다이제스트 없음</div>'

    # 하단: 카페 상세글
    cafe_col_html = "".join(issue_item_html(n) for n in cafe_items) or '<div class="issue-empty">수집된 카페글 없음</div>'

    # 테이블 행
    table_rows = ""
    for r in reversed(data):
        result_cls   = {"승": "win", "패": "lose", "무": "draw", "취소": "cancel"}.get(r["result"], "")
        result_txt   = r["result"] or "-"
        crowd        = r.get("crowd", 0)
        occupancy    = f"{crowd / STADIUM_CAPACITY * 100:.1f}%" if crowd and r["home_away"] == "홈" else "-"
        crowd_txt    = fmt(crowd) if crowd else "-"
        table_rows += f"""
        <tr data-month="{r['date'][:7]}">
          <td>{r["date"]}</td>
          <td>{r["home_away"] or "-"}</td>
          <td>{r["opponent"] or "-"}</td>
          <td><span class="badge {result_cls}">{result_txt}</span></td>
          <td class="num">{crowd_txt}</td>
          <td class="num">{occupancy}</td>
          <td class="num">{fmt(r["off"])}</td>
          <td class="num">{fmt(r["on"])}</td>
          <td class="num bold">{fmt(r["total"])}</td>
        </tr>"""

    updated_at = datetime.now().strftime("%Y.%m.%d %H:%M")

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>삼성 라이온즈 매출 대시보드</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>
<style>
  @font-face {{
    font-family: 'KBO-Dia-Gothic';
    src: url('https://cdn.jsdelivr.net/gh/projectnoonnu/noonfonts_2304-2@1.0/KBO-Dia-Gothic_light.woff') format('woff');
    font-weight: 300; font-style: normal; font-display: swap;
  }}
  @font-face {{
    font-family: 'KBO-Dia-Gothic';
    src: url('https://cdn.jsdelivr.net/gh/projectnoonnu/noonfonts_2304-2@1.0/KBO-Dia-Gothic_medium.woff') format('woff');
    font-weight: 500; font-style: normal; font-display: swap;
  }}
  @font-face {{
    font-family: 'KBO-Dia-Gothic';
    src: url('https://cdn.jsdelivr.net/gh/projectnoonnu/noonfonts_2304-2@1.0/KBO-Dia-Gothic_bold.woff') format('woff');
    font-weight: 700; font-style: normal; font-display: swap;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'KBO-Dia-Gothic', 'Malgun Gothic', sans-serif; font-weight: 500; background: #f0f2f5; color: #1a1a2e; }}

  /* 헤더 */
  .header {{
    background: linear-gradient(135deg, #002D72 0%, #C8102E 100%);
    color: white; padding: 24px 32px;
    display: flex; justify-content: space-between; align-items: center;
  }}
  .header h1 {{ font-size: 22px; font-weight: 700; letter-spacing: -0.5px; }}
  .header .updated {{ font-size: 12px; opacity: 0.8; }}

  /* KPI */
  .kpi-row {{ display: flex; gap: 16px; padding: 24px 32px 0; flex-wrap: wrap; }}
  .kpi {{
    flex: 1; min-width: 160px; background: white;
    border-radius: 12px; padding: 20px 24px;
    box-shadow: 0 2px 8px rgba(0,0,0,.07);
  }}
  .kpi .label {{ font-size: 12px; color: #888; margin-bottom: 8px; }}
  .kpi .value {{ font-size: 24px; font-weight: 700; color: #002D72; }}
  .kpi .value.red {{ color: #C8102E; }}
  .kpi .value .unit {{ font-size: 13px; font-weight: 400; color: #888; margin-left: 3px; }}
  .kpi .sub   {{ font-size: 11px; color: #aaa; margin-top: 4px; }}

  /* 차트 그리드 */
  .charts {{ display: grid; grid-template-columns: 2fr 1fr 1fr; grid-auto-rows: 320px; gap: 16px; padding: 20px 32px; }}
  .chart-card {{
    background: white; border-radius: 12px; padding: 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,.07);
  }}
  .chart-card h3 {{ font-size: 13px; color: #555; margin-bottom: 16px; font-weight: 600; }}
  .chart-card.side {{
    display: flex; flex-direction: column; padding-bottom: 12px;
  }}
  .chart-card.side .chart-wrap {{
    flex: 1; min-height: 0; position: relative;
  }}

  /* 테이블 */
  .table-section {{ padding: 0 32px 20px; }}
  .table-card {{
    background: white; border-radius: 12px; padding: 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,.07); overflow-x: auto;
  }}
  .table-card h3 {{ font-size: 13px; color: #555; margin-bottom: 16px; font-weight: 600; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  th {{ background: #f8f9fa; padding: 10px 12px; text-align: left; color: #666; font-weight: 600; border-bottom: 2px solid #eee; }}
  td {{ padding: 10px 12px; border-bottom: 1px solid #f0f0f0; }}
  td.num {{ text-align: right; font-variant-numeric: tabular-nums; }}
  td.bold {{ font-weight: 700; color: #002D72; }}
  tr:hover {{ background: #fafbff; }}

  /* 이슈 섹션 */
  .issue-section {{ padding: 0 32px 32px; }}
  .issue-section-title {{
    font-size: 14px; font-weight: 700; color: #333;
    margin-bottom: 12px; padding-left: 4px;
  }}
  .issue-layout {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; align-items: start; }}
  .issue-col-left {{ display: flex; flex-direction: column; }}
  .issue-col-right {{ display: flex; flex-direction: column; gap: 16px; }}
  .issue-box {{
    background: white; border-radius: 12px; padding: 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,.07); overflow-y: auto;
  }}
  .issue-box h4 {{
    font-size: 12px; font-weight: 700; color: #002D72;
    margin-bottom: 12px; padding-bottom: 6px; border-bottom: 2px solid #002D72;
  }}
  .issue-box.cafe-digest h4 {{ color: #2e7d32; border-color: #2e7d32; }}
  .issue-box.cafe-posts  h4 {{ color: #2e7d32; border-color: #2e7d32; }}
  .issue-item {{ margin-bottom: 14px; padding-bottom: 14px; border-bottom: 1px solid #f0f0f0; }}
  .issue-item:last-child {{ margin-bottom: 0; padding-bottom: 0; border-bottom: none; }}
  .issue-meta {{ font-size: 10px; color: #aaa; margin-bottom: 3px; }}
  .issue-views {{ font-size: 10px; color: #888; }}
  .issue-title {{ font-size: 13px; color: #222; line-height: 1.5; font-weight: 500; }}
  .issue-title a {{ color: #222; text-decoration: none; }}
  .issue-title a:hover {{ color: #002D72; text-decoration: underline; }}
  .issue-desc {{ font-size: 11px; color: #777; margin-top: 4px; line-height: 1.5; }}
  .issue-empty {{ font-size: 12px; color: #bbb; text-align: center; padding: 20px 0; }}
  .digest-text {{ font-size: 12px; color: #333; line-height: 1.8; white-space: pre-line; }}

  .filter-bar {{ display: flex; gap: 8px; flex-wrap: wrap; }}
  .filter-btn {{
    padding: 6px 14px; border: 1.5px solid #002D72; border-radius: 20px;
    background: white; color: #002D72; font-size: 12px; font-weight: 600;
    cursor: pointer; transition: all 0.15s;
  }}
  .filter-btn.active {{ background: #002D72; color: white; }}
  .filter-btn:hover:not(.active) {{ background: #e8edf5; }}
  .summary-total {{ background: #f0f4ff; font-weight: 700; }}
  #excelDownloadBtn {{
    padding: 7px 16px; background: #217346; color: white; border: none;
    border-radius: 8px; font-size: 12px; font-weight: 600; cursor: pointer;
  }}
  #excelDownloadBtn:hover {{ background: #185c37; }}

  /* 페이지네이션 */
  .pagination {{ display: flex; gap: 6px; justify-content: center; padding: 16px 0 4px; flex-wrap: wrap; }}
  .page-btn {{
    padding: 5px 12px; border: 1.5px solid #ddd; border-radius: 6px;
    background: white; color: #444; font-size: 12px; font-weight: 600;
    cursor: pointer; transition: all 0.15s; min-width: 36px;
  }}
  .page-btn.active {{ background: #002D72; color: white; border-color: #002D72; }}
  .page-btn:hover:not(.active) {{ background: #f0f4ff; border-color: #002D72; }}

  .badge {{ display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 700; }}
  .badge.win    {{ background: #e8f5e9; color: #2e7d32; }}
  .badge.lose   {{ background: #ffebee; color: #c62828; }}
  .badge.draw   {{ background: #e3f2fd; color: #1565c0; }}
  .badge.cancel {{ background: #f5f5f5; color: #757575; }}

  /* 특이사항 모달 */
  .modal-overlay {{
    display: none; position: fixed; inset: 0; z-index: 1000;
    background: rgba(0,0,0,.45); align-items: center; justify-content: center;
  }}
  .modal-overlay.open {{ display: flex; }}
  .modal {{
    background: white; border-radius: 16px; padding: 28px 32px;
    width: 420px; max-width: 90vw; box-shadow: 0 8px 32px rgba(0,0,0,.2);
  }}
  .modal h3 {{ font-size: 16px; font-weight: 700; color: #002D72; margin-bottom: 20px; }}
  .modal label {{ font-size: 12px; color: #666; font-weight: 600; display: block; margin-bottom: 6px; }}
  .modal input, .modal textarea {{
    width: 100%; border: 1.5px solid #ddd; border-radius: 8px;
    padding: 10px 12px; font-size: 13px; margin-bottom: 16px;
    font-family: inherit; outline: none;
  }}
  .modal input:focus, .modal textarea:focus {{ border-color: #002D72; }}
  .modal textarea {{ height: 80px; resize: vertical; }}
  .modal-btns {{ display: flex; gap: 10px; justify-content: flex-end; }}
  .modal-btns button {{
    padding: 9px 20px; border-radius: 8px; font-size: 13px;
    font-weight: 600; cursor: pointer; border: none;
  }}
  .btn-cancel {{ background: #f0f0f0; color: #555; }}
  .btn-submit {{ background: #002D72; color: white; }}
  .btn-submit:hover {{ background: #001a4a; }}
  .btn-submit:disabled {{ background: #aaa; cursor: not-allowed; }}
  .modal-status {{ font-size: 12px; margin-top: 10px; text-align: center; min-height: 18px; }}

  /* 상품별 매출 섹션 */
  .product-section {{ padding: 0 32px 20px; }}
  #productExcelBtn {{
    padding: 7px 16px; background: #217346; color: white; border: none;
    border-radius: 8px; font-size: 12px; font-weight: 600; cursor: pointer;
  }}
  #productExcelBtn:hover {{ background: #185c37; }}
  .product-total-row {{ background: #f0f4ff; font-weight: 700; }}
  .product-controls {{ display:flex; gap:10px; flex-wrap:wrap; align-items:center; margin-bottom:12px; }}
  .product-controls input[type="date"], .product-controls input[type="text"] {{
    padding:6px 10px; border:1.5px solid #ddd; border-radius:8px; font-size:12px;
    font-family:inherit; outline:none;
  }}
  .product-controls input:focus {{ border-color:#002D72; }}
  .product-controls .ctrl-btn {{
    padding:6px 14px; border-radius:8px; font-size:12px; font-weight:600; cursor:pointer; border:none;
  }}
  .product-controls .ctrl-btn-primary {{ background:#002D72; color:white; }}
  .product-controls .ctrl-btn-primary:hover {{ background:#001a4a; }}
  .product-controls .ctrl-btn-reset {{ background:#f0f0f0; color:#555; }}
  .product-controls .ctrl-btn-reset:hover {{ background:#e0e0e0; }}
  .product-guide {{ font-size:11px; color:#888; margin-bottom:12px; }}
  .drilldown-section {{ padding:0 32px 20px; }}
  .drilldown-charts {{ display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-top:16px; }}
  .drilldown-chart-wrap {{ height:220px; position:relative; }}

  @media (max-width: 900px) {{
    .charts {{ grid-template-columns: 1fr; }}
    .kpi-row, .charts, .table-section, .product-section, .issue-section {{ padding-left: 16px; padding-right: 16px; }}
    .issue-layout {{ grid-template-columns: 1fr; }}
  }}
</style>
</head>
<body>

<div class="header">
  <h1>⚾ 삼성 라이온즈 매출 대시보드</h1>
  <span class="updated">최종 업데이트: {updated_at}</span>
</div>

<div class="kpi-row">
  <div class="kpi">
    <div class="label">OFF 거래액 (누계)</div>
    <div class="value"><span id="kpi-off">{fmt(total_off)}</span><span class="unit">원</span></div>
    <div class="sub">{date_range_label}</div>
  </div>
  <div class="kpi">
    <div class="label">ON 거래액 (누계)</div>
    <div class="value"><span id="kpi-on">{fmt(total_on)}</span><span class="unit">원</span></div>
    <div class="sub">{date_range_label}</div>
  </div>
  <div class="kpi">
    <div class="label">총 거래액 (누계)</div>
    <div class="value red"><span id="kpi-total">{fmt(total_total)}</span><span class="unit">원</span></div>
    <div class="sub">{date_range_label}</div>
  </div>
  <div class="kpi">
    <div class="label">경기 수</div>
    <div class="value">{len(game_days)}</div>
    <div class="sub">홈 {len(home_rows)} · 어웨이 {len(away_rows)}</div>
  </div>
  <div class="kpi">
    <div class="label">승률</div>
    <div class="value">{win_rate}</div>
    <div class="sub">{wins}승 {losses}패</div>
  </div>
  <div class="kpi">
    <div class="label">홈 평균 관중</div>
    <div class="value">{fmt(avg_crowd)}<span class="unit">명</span></div>
    <div class="sub">평균 점유율 {avg_occupancy} (수용 {STADIUM_CAPACITY:,}명)</div>
  </div>
</div>

<div class="charts">
  <div class="chart-card">
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:16px">
      <h3 style="margin-bottom:0">날짜별 거래액 추이</h3>
      <button onclick="openNoteModal()" title="특이사항 입력"
        style="background:none;border:1.5px solid #ddd;border-radius:8px;padding:4px 10px;cursor:pointer;font-size:13px;color:#555;white-space:nowrap;">
        ✏️ 특이사항
      </button>
    </div>
    <div style="position:relative;width:100%;height:calc(100% - 48px)">
      <canvas id="trendChart"></canvas>
    </div>
  </div>
  <div class="chart-card side">
    <h3>홈 / 어웨이 평균 거래액</h3>
    <div class="chart-wrap">
      <canvas id="haChart"></canvas>
    </div>
  </div>
  <div class="chart-card side">
    <h3>경기 결과별 평균 거래액</h3>
    <div class="chart-wrap">
      <canvas id="resultChart"></canvas>
    </div>
  </div>
</div>

<div class="table-section">
  <!-- 월별 요약 -->
  <div class="table-card" style="margin-bottom:16px">
    <h3>월별 합계</h3>
    <table>
      <thead>
        <tr>
          <th>월</th>
          <th style="text-align:right">OFF거래액</th>
          <th style="text-align:right">ON거래액</th>
          <th style="text-align:right">합계</th>
        </tr>
      </thead>
      <tbody>{summary_rows_html}</tbody>
    </table>
  </div>
  <!-- 전체 데이터 -->
  <div class="table-card">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px">
      <h3 style="margin-bottom:0">전체 데이터</h3>
      <button id="excelDownloadBtn" onclick="downloadExcel()">📥 엑셀 다운로드</button>
    </div>
    <!-- 월별 필터 -->
    <div class="filter-bar" style="margin-bottom:16px">
      {filter_btns_html}
    </div>
    <table id="dataTable">
      <thead>
        <tr>
          <th>날짜</th><th>홈/어웨이</th><th>상대팀</th><th>결과</th>
          <th style="text-align:right">관중수</th>
          <th style="text-align:right">점유율</th>
          <th style="text-align:right">OFF거래액</th>
          <th style="text-align:right">ON거래액</th>
          <th style="text-align:right">합계</th>
        </tr>
      </thead>
      <tbody>{table_rows}</tbody>
    </table>
    <div class="pagination" id="pagination"></div>
  </div>
</div>

<div class="product-section">
  <div class="table-card">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
      <div>
        <h3 style="margin-bottom:4px">상품별 누적 판매 실적 (오프라인)</h3>
        <span id="productRangeLabel" style="font-size:11px;color:#aaa"></span>
      </div>
      <button id="productExcelBtn" onclick="downloadProductExcel()">📥 엑셀 다운로드</button>
    </div>
    <div class="product-controls">
      <input type="date" id="prodStartDate" title="시작일">
      <span style="font-size:12px;color:#888">~</span>
      <input type="date" id="prodEndDate" title="종료일">
      <input type="text" id="prodNameSearch" placeholder="상품명 검색" style="width:160px">
      <button class="ctrl-btn ctrl-btn-primary" onclick="applyProductFilter()">조회</button>
      <button class="ctrl-btn ctrl-btn-reset" onclick="resetProductFilter()">초기화</button>
    </div>
    <p class="product-guide">💡 기본 화면은 전체 기간 누적 매출 상위 10개 상품입니다. 기간·상품명 필터 후 조회하거나, 엑셀 다운로드로 날짜별 전체 상품 데이터를 확인하세요.</p>
    <table id="productTable">
      <thead>
        <tr>
          <th>상품코드</th><th>상품명</th><th>칼라</th><th>사이즈</th>
          <th style="text-align:right">판매단가</th>
          <th style="text-align:right">판매수량(누적)</th>
          <th style="text-align:right">실판매금액(누적)</th>
          <th style="text-align:center">트렌드</th>
        </tr>
      </thead>
      <tbody id="productTbody"></tbody>
    </table>
    <div class="pagination" id="productPagination"></div>
  </div>
</div>

<div class="drilldown-section" id="drilldownSection" style="display:none">
  <div class="table-card">
    <div style="display:flex;justify-content:space-between;align-items:center">
      <div>
        <h3 id="drilldownTitle" style="margin-bottom:4px"></h3>
        <span id="drilldownSub" style="font-size:11px;color:#aaa"></span>
      </div>
      <button onclick="closeDrilldown()" style="background:#f0f0f0;border:none;border-radius:8px;padding:6px 14px;cursor:pointer;font-size:12px;font-weight:600;">✕ 닫기</button>
    </div>
    <div class="drilldown-charts">
      <div>
        <div style="font-size:12px;color:#555;font-weight:600;margin-bottom:8px">일별 실판매금액</div>
        <div class="drilldown-chart-wrap"><canvas id="drilldownAmountChart"></canvas></div>
      </div>
      <div>
        <div style="font-size:12px;color:#555;font-weight:600;margin-bottom:8px">일별 판매수량</div>
        <div class="drilldown-chart-wrap"><canvas id="drilldownQtyChart"></canvas></div>
      </div>
    </div>
  </div>
</div>

<div class="issue-section">
  <div class="issue-section-title">최근 7일간 주요이슈</div>
  <div class="issue-layout">
    <div class="issue-col-left">
      <div class="issue-box">
        <h4>뉴스</h4>
        {news_col_html}
      </div>
    </div>
    <div class="issue-col-right">
      <div class="issue-box cafe-digest">
        <h4>카페 트렌드 다이제스트</h4>
        {digest_col_html}
      </div>
      <div class="issue-box cafe-posts">
        <h4>카페 인기글 TOP 10 (사자사랑방)</h4>
        {cafe_col_html}
      </div>
    </div>
  </div>
</div>

<script>
// ── 월별 필터 & 페이지네이션 ──
const allData = {json.dumps(data)};
const ROWS_PER_PAGE = 15;
let currentMonth = 'all';
let currentPage  = 1;

document.querySelectorAll('.filter-btn').forEach(btn => {{
  btn.addEventListener('click', function() {{
    document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
    this.classList.add('active');
    currentMonth = this.dataset.month;
    currentPage  = 1;
    updateKPI(currentMonth);
    renderTable();
  }});
}});

function getFilteredRows() {{
  const allRows = Array.from(document.querySelectorAll('#dataTable tbody tr'));
  return currentMonth === 'all'
    ? allRows
    : allRows.filter(row => row.dataset.month === currentMonth);
}}

function renderTable() {{
  const allRows     = Array.from(document.querySelectorAll('#dataTable tbody tr'));
  const filtered    = getFilteredRows();
  const totalPages  = Math.ceil(filtered.length / ROWS_PER_PAGE);
  const start       = (currentPage - 1) * ROWS_PER_PAGE;
  const pageRows    = filtered.slice(start, start + ROWS_PER_PAGE);

  allRows.forEach(row => row.style.display = 'none');
  pageRows.forEach(row => row.style.display = '');

  renderPagination(totalPages);
}}

function renderPagination(totalPages) {{
  const container = document.getElementById('pagination');
  if (totalPages <= 1) {{ container.innerHTML = ''; return; }}
  let html = '';
  for (let i = 1; i <= totalPages; i++) {{
    html += `<button class="page-btn ${{i === currentPage ? 'active' : ''}}" onclick="goToPage(${{i}})">${{i}}</button>`;
  }}
  container.innerHTML = html;
}}

function goToPage(page) {{
  currentPage = page;
  renderTable();
  document.querySelector('.table-card:last-of-type').scrollIntoView({{behavior:'smooth', block:'start'}});
}}

function updateKPI(month) {{
  const filtered = month === 'all' ? allData : allData.filter(r => r.date.startsWith(month));
  const off   = filtered.reduce((s, r) => s + r.off, 0);
  const on    = filtered.reduce((s, r) => s + r.on, 0);
  const total = off + on;
  const fmt   = n => (n && n !== 0) ? n.toLocaleString('ko-KR') : '-';
  document.getElementById('kpi-off').textContent   = fmt(off);
  document.getElementById('kpi-on').textContent    = fmt(on);
  document.getElementById('kpi-total').textContent = fmt(total);
}}

// ── 엑셀 다운로드 (현재 필터 기준 전체 데이터) ──
function downloadExcel() {{
  const filtered = getFilteredRows();
  const headers  = ['날짜','홈/어웨이','상대팀','결과','관중수','점유율','OFF거래액','ON거래액','합계'];
  const csvData  = [headers];
  filtered.forEach(row => {{
    const cells = row.querySelectorAll('td');
    csvData.push(Array.from(cells).map(td => td.textContent.trim()));
  }});
  const ws = XLSX.utils.aoa_to_sheet(csvData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '매출데이터');
  XLSX.writeFile(wb, '삼성라이온즈_매출데이터.xlsx');
}}

// 초기 렌더링
renderTable();

const OFF = '#002D72', ON = '#C8102E', TOT = '#6c757d';
const alpha = (c, a) => c + Math.round(a*255).toString(16).padStart(2,'0');

// ── 상품별 매출 ──
const rawProductData = {raw_products_json};
const PRODUCT_ROWS_PER_PAGE = 10;
let productPage = 1;
let currentProductRows = [];
let drilldownAmtChart = null;
let drilldownQtyChart = null;

// 날짜 비교용: "2026.03.16" → "2026-03-16"
function dotToHyphen(d) {{ return d ? d.replace(/\./g, '-') : ''; }}
function hyphenToDot(d) {{ return d ? d.replace(/-/g, '.') : ''; }}

function aggregateProducts(rows) {{
  const agg = {{}};
  rows.forEach(r => {{
    const key = r.code + '|' + r.name + '|' + r.color + '|' + r.size;
    if (!agg[key]) agg[key] = {{code:r.code, name:r.name, color:r.color, size:r.size, price:r.price, qty:0, amount:0}};
    agg[key].qty    += r.qty;
    agg[key].amount += r.amount;
  }});
  return Object.values(agg).sort((a,b) => b.amount - a.amount);
}}

function getFilteredRaw() {{
  const start = document.getElementById('prodStartDate').value;
  const end   = document.getElementById('prodEndDate').value;
  const name  = document.getElementById('prodNameSearch').value.trim();
  return rawProductData.filter(r => {{
    if (start && dotToHyphen(r.date) < start) return false;
    if (end   && dotToHyphen(r.date) > end)   return false;
    if (name  && !r.name.includes(name))       return false;
    return true;
  }});
}}

function applyProductFilter() {{
  const filtered = getFilteredRaw();
  currentProductRows = aggregateProducts(filtered);
  productPage = 1;

  // 기간 레이블 업데이트
  const start = document.getElementById('prodStartDate').value;
  const end   = document.getElementById('prodEndDate').value;
  const name  = document.getElementById('prodNameSearch').value.trim();
  let label = '';
  if (start || end) label += (hyphenToDot(start)||'전체') + ' ~ ' + (hyphenToDot(end)||'전체');
  if (name) label += (label ? ' / ' : '') + '검색: ' + name;
  document.getElementById('productRangeLabel').textContent = label || '';

  renderProductTable();
}}

function resetProductFilter() {{
  document.getElementById('prodStartDate').value = '';
  document.getElementById('prodEndDate').value   = '';
  document.getElementById('prodNameSearch').value = '';
  currentProductRows = aggregateProducts(rawProductData).slice(0, 10);
  productPage = 1;
  document.getElementById('productRangeLabel').textContent = '누적 매출 상위 10개 상품';
  renderProductTable();
}}

function renderProductTable() {{
  const tbody = document.getElementById('productTbody');
  const total      = currentProductRows.length;
  const totalPages = Math.ceil(total / PRODUCT_ROWS_PER_PAGE);
  const start      = (productPage - 1) * PRODUCT_ROWS_PER_PAGE;
  const pageRows   = currentProductRows.slice(start, start + PRODUCT_ROWS_PER_PAGE);

  // 합계
  const totalQty    = currentProductRows.reduce((s,p) => s + p.qty, 0);
  const totalAmount = currentProductRows.reduce((s,p) => s + p.amount, 0);
  const fmt = n => n ? n.toLocaleString('ko-KR') : '-';

  let html = '';
  pageRows.forEach(p => {{
    html += `<tr style="cursor:pointer" onclick="showDrilldown('${{p.code}}','${{p.name}}','${{p.color}}','${{p.size}}')">
      <td>${{p.code}}</td><td>${{p.name}}</td><td>${{p.color}}</td><td>${{p.size}}</td>
      <td class="num">${{fmt(p.price)}}</td>
      <td class="num">${{p.qty.toLocaleString()}}</td>
      <td class="num bold">${{fmt(p.amount)}}</td>
      <td style="text-align:center;font-size:13px;color:#002D72">📈</td>
    </tr>`;
  }});
  // 합계 행
  html += `<tr class="product-total-row">
    <td colspan="5">합계 (전체 ${{total}}개 SKU)</td>
    <td class="num">${{totalQty.toLocaleString()}}</td>
    <td class="num bold">${{fmt(totalAmount)}}</td>
    <td></td>
  </tr>`;
  tbody.innerHTML = html;

  // 페이지네이션
  const container = document.getElementById('productPagination');
  if (totalPages <= 1) {{ container.innerHTML = ''; return; }}
  const maxBtn = 10;
  let startBtn = Math.max(1, productPage - Math.floor(maxBtn/2));
  let endBtn   = Math.min(totalPages, startBtn + maxBtn - 1);
  if (endBtn - startBtn < maxBtn-1) startBtn = Math.max(1, endBtn - maxBtn + 1);
  let phtml = '';
  if (startBtn > 1) phtml += `<button class="page-btn" onclick="goProductPage(1)">1</button><span style="padding:0 4px;color:#aaa">…</span>`;
  for (let i = startBtn; i <= endBtn; i++) {{
    phtml += `<button class="page-btn ${{i===productPage?'active':''}}" onclick="goProductPage(${{i}})">${{i}}</button>`;
  }}
  if (endBtn < totalPages) phtml += `<span style="padding:0 4px;color:#aaa">…</span><button class="page-btn" onclick="goProductPage(${{totalPages}})">${{totalPages}}</button>`;
  container.innerHTML = phtml;
}}

function goProductPage(page) {{
  productPage = page;
  renderProductTable();
  document.querySelector('.product-section').scrollIntoView({{behavior:'smooth', block:'start'}});
}}

// ── 상품 드릴다운 차트 ──
function showDrilldown(code, name, color, size) {{
  const filtered = getFilteredRaw().filter(r =>
    r.code === code && r.name === name && r.color === color && r.size === size
  );
  const byDate = {{}};
  filtered.forEach(r => {{
    if (!byDate[r.date]) byDate[r.date] = {{qty:0, amount:0}};
    byDate[r.date].qty    += r.qty;
    byDate[r.date].amount += r.amount;
  }});
  const dates   = Object.keys(byDate).sort();
  const amounts = dates.map(d => byDate[d].amount);
  const qtys    = dates.map(d => byDate[d].qty);
  const labels  = dates.map(d => d.slice(5));

  document.getElementById('drilldownTitle').textContent = `${{name}} (${{color}} / ${{size}})`;
  document.getElementById('drilldownSub').textContent   = `${{code}} · 조회 기간 내 ${{dates.length}}일 판매`;
  document.getElementById('drilldownSection').style.display = 'block';
  document.getElementById('drilldownSection').scrollIntoView({{behavior:'smooth', block:'start'}});

  if (drilldownAmtChart) drilldownAmtChart.destroy();
  if (drilldownQtyChart) drilldownQtyChart.destroy();

  drilldownAmtChart = new Chart(document.getElementById('drilldownAmountChart'), {{
    type: 'line',
    data: {{
      labels,
      datasets: [{{ label:'실판매금액', data:amounts,
        borderColor:'#002D72', backgroundColor:'rgba(0,45,114,0.1)',
        fill:true, tension:0.3, pointRadius:3 }}]
    }},
    options: {{ responsive:true, maintainAspectRatio:false,
      scales:{{ y:{{ ticks:{{ callback: v => v>=1e6?(v/1e6).toFixed(1)+'M':v.toLocaleString() }} }} }} }}
  }});

  drilldownQtyChart = new Chart(document.getElementById('drilldownQtyChart'), {{
    type: 'bar',
    data: {{
      labels,
      datasets: [{{ label:'판매수량', data:qtys, backgroundColor:'rgba(200,16,46,0.7)' }}]
    }},
    options: {{ responsive:true, maintainAspectRatio:false }}
  }});
}}

function closeDrilldown() {{
  document.getElementById('drilldownSection').style.display = 'none';
  if (drilldownAmtChart) {{ drilldownAmtChart.destroy(); drilldownAmtChart = null; }}
  if (drilldownQtyChart) {{ drilldownQtyChart.destroy(); drilldownQtyChart = null; }}
}}

// ── 상품별 매출 엑셀 다운로드 (날짜별 전체) ──
function downloadProductExcel() {{
  const headers = ['판매일자','상품코드','상품명','칼라','사이즈','판매단가','판매수량','실판매금액'];
  const rows = rawProductData
    .slice()
    .sort((a,b) => a.date.localeCompare(b.date) || a.name.localeCompare(b.name))
    .map(r => [r.date, r.code, r.name, r.color, r.size, r.price, r.qty, r.amount]);
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '상품별매출_날짜별');
  XLSX.writeFile(wb, '삼성라이온즈_상품별매출_날짜별전체.xlsx');
}}

// 초기 렌더링: 상위 10개
currentProductRows = aggregateProducts(rawProductData).slice(0, 10);
document.getElementById('productRangeLabel').textContent = '누적 매출 상위 10개 상품';
renderProductTable();

// ── 특이사항 노트 플러그인 ──
const chartNotes = {json.dumps(chart_notes)};
const notePlugin = {{
  id: 'notePlugin',
  afterDatasetsDraw(chart) {{
    const ctx = chart.ctx;
    const numDatasets = chart.data.datasets.length;
    const firstMeta = chart.getDatasetMeta(0);

    firstMeta.data.forEach((pt, i) => {{
      const note = chartNotes[i];
      if (!note) return;

      // 모든 데이터셋 중 가장 높은 점(y값 최소) 찾기
      let minY = pt.y;
      for (let d = 1; d < numDatasets; d++) {{
        const m = chart.getDatasetMeta(d);
        if (m.data[i] && !m.hidden) {{
          minY = Math.min(minY, m.data[i].y);
        }}
      }}

      ctx.save();
      ctx.font = 'bold 10px sans-serif';
      ctx.fillStyle = '#C8102E';
      ctx.textAlign = 'center';
      const text = note.length > 16 ? note.substring(0, 15) + '…' : note;
      ctx.translate(pt.x, minY - 16);
      ctx.rotate(-Math.PI / 8);
      ctx.fillText(text, 0, 0);
      ctx.restore();
    }});
  }}
}};

// ── 추이 차트 ──
new Chart(document.getElementById('trendChart'), {{
  type: 'line',
  plugins: [notePlugin],
  data: {{
    labels: {json.dumps(chart_dates)},
    datasets: [
      {{
        label: 'OFF거래액', data: {json.dumps(chart_off)},
        borderColor: OFF, backgroundColor: alpha(OFF, .15),
        borderWidth: 2, pointRadius: 3, pointHoverRadius: 5, fill: true, tension: 0.3,
      }},
      {{
        label: 'ON거래액', data: {json.dumps(chart_on)},
        borderColor: ON, backgroundColor: alpha(ON, .1),
        borderWidth: 2, pointRadius: 3, pointHoverRadius: 5, fill: true, tension: 0.3,
      }},
    ]
  }},
  options: {{
    responsive: true,
    maintainAspectRatio: false,
    layout: {{ padding: {{ top: 30 }} }},
    plugins: {{ legend: {{ position: 'top' }} }},
    scales: {{
      x: {{ ticks: {{ maxTicksLimit: 12 }} }},
      y: {{ ticks: {{ callback: v => v >= 1e6 ? (v/1e6).toFixed(1)+'M' : v.toLocaleString() }} }}
    }}
  }}
}});

// 바 상단 수치 (백만 단위 반올림) 인라인 플러그인
const topLabelPlugin = {{
  id: 'topLabel',
  afterDatasetsDraw(chart) {{
    const ctx = chart.ctx;
    chart.data.datasets.forEach((ds, di) => {{
      chart.getDatasetMeta(di).data.forEach((bar, i) => {{
        const v = ds.data[i];
        if (!v) return;
        ctx.save();
        ctx.font = 'bold 11px sans-serif';
        ctx.fillStyle = '#444';
        ctx.textAlign = 'center';
        ctx.fillText(Math.round(v / 1e6) + '백만', bar.x, bar.y - 5);
        ctx.restore();
      }});
    }});
  }}
}};

const sideOpts = () => ({{
  responsive: true,
  maintainAspectRatio: false,
  plugins: {{ legend: {{ position: 'bottom' }} }},
  scales: {{
    y: {{
      min: 0,
      ticks: {{ display: false }},
      grid: {{ display: false }},
      border: {{ display: false }},
    }}
  }},
  layout: {{ padding: {{ top: 24, bottom: 0 }} }}
}});

// ── 홈/어웨이 월별 ──
new Chart(document.getElementById('haChart'), {{
  type: 'bar',
  plugins: [topLabelPlugin],
  data: {{
    labels: {json.dumps(ha_labels)},
    datasets: [
      {{ label: 'OFF거래액', data: {json.dumps(ha_off_data)}, backgroundColor: alpha(OFF, .8) }},
      {{ label: 'ON거래액',  data: {json.dumps(ha_on_data)},  backgroundColor: alpha(ON,  .8) }},
    ]
  }},
  options: sideOpts(),
}});

// ── 결과별 월별 ──
new Chart(document.getElementById('resultChart'), {{
  type: 'bar',
  plugins: [topLabelPlugin],
  data: {{
    labels: {json.dumps(res_labels)},
    datasets: [
      {{ label: 'OFF거래액', data: {json.dumps(res_off_data)}, backgroundColor: alpha(OFF, .8) }},
      {{ label: 'ON거래액',  data: {json.dumps(res_on_data)},  backgroundColor: alpha(ON,  .8) }},
    ]
  }},
  options: sideOpts(),
}});

// ── 특이사항 입력 모달 ──
const WORKER_URL = 'https://lions-note.freelywind222.workers.dev';

function openNoteModal() {{
  const today = new Date();
  const y = today.getFullYear();
  const m = String(today.getMonth()+1).padStart(2,'0');
  const d = String(today.getDate()).padStart(2,'0');
  document.getElementById('noteDate').value = `${{y}}-${{m}}-${{d}}`;
  document.getElementById('noteText').value = '';
  document.getElementById('noteStatus').textContent = '';
  document.getElementById('noteSubmitBtn').disabled = false;
  document.getElementById('noteModal').classList.add('open');
}}

function closeNoteModal() {{
  document.getElementById('noteModal').classList.remove('open');
}}

document.getElementById('noteModal').addEventListener('click', function(e) {{
  if (e.target === this) closeNoteModal();
}});

async function submitNote() {{
  const rawDate = document.getElementById('noteDate').value;
  const date = rawDate ? rawDate.replace(/-/g, '.') : '';  // YYYY-MM-DD → YYYY.MM.DD
  const note = document.getElementById('noteText').value.trim();
  const status = document.getElementById('noteStatus');
  const btn = document.getElementById('noteSubmitBtn');

  if (!date || !note) {{ status.textContent = '날짜와 내용을 모두 입력해주세요.'; return; }}

  btn.disabled = true;
  status.style.color = '#666';
  status.textContent = '업데이트 중... (약 1~2분 소요)';

  try {{
    const res = await fetch(WORKER_URL, {{
      method: 'POST',
      headers: {{ 'Content-Type': 'application/json' }},
      body: JSON.stringify({{ date, note }}),
    }});
    if (res.ok) {{
      status.style.color = '#2e7d32';
      status.textContent = '✓ 업데이트 요청 완료! 대시보드가 곧 반영됩니다.';
      setTimeout(closeNoteModal, 2500);
    }} else {{
      throw new Error(await res.text());
    }}
  }} catch(e) {{
    status.style.color = '#c62828';
    status.textContent = '오류: ' + e.message;
    btn.disabled = false;
  }}
}}
</script>

<!-- 특이사항 모달 -->
<div class="modal-overlay" id="noteModal">
  <div class="modal">
    <h3>✏️ 특이사항 입력</h3>
    <label>날짜</label>
    <input type="date" id="noteDate">
    <label>특이사항</label>
    <textarea id="noteText" placeholder="예) 홈 개막전, 유니폼 판매 행사"></textarea>
    <div class="modal-btns">
      <button class="btn-cancel" onclick="closeNoteModal()">취소</button>
      <button class="btn-submit" id="noteSubmitBtn" onclick="submitNote()">저장</button>
    </div>
    <div class="modal-status" id="noteStatus"></div>
  </div>
</div>

</body>
</html>"""


def main():
    print("데이터 조회 중...")
    data, news, digest, raw_products = fetch_data()
    print(f"병합된 데이터: {len(data)}일 / 뉴스이슈: {len(news)}건 / 상품데이터: {len(raw_products)}행")

    os.makedirs("dashboard", exist_ok=True)
    html = build_html(data, news, digest, raw_products)
    with open("dashboard/index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("dashboard/index.html 생성 완료")


if __name__ == "__main__":
    main()
