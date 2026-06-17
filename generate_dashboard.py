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

    def _int(v):
        try: return int(float(str(v).replace(",", "")))
        except: return 0

    def _barcode(v):
        s = str(v).strip()
        if not s or s in ("", "None", "nan"): return ""
        try: return str(int(float(s)))  # "8804775462283.0" → "8804775462283"
        except: return s

    def _norm_date(d):
        # 2026.05.20 → 2026-05-20 으로 통일
        return str(d).strip().replace(".", "-")

    # ── 상품별매출(off) ───────────────────────────────────
    raw_products_off = []
    try:
        ws_off  = sh.worksheet("상품별매출(off)")
        off_raw = ws_off.get_all_records()
        for row in off_raw:
            date    = _norm_date(row.get("판매일자", "") or "")
            barcode = _barcode(row.get("추가바코드1",""))
            name    = str(row.get("상품명",     "") or "").strip()
            color   = str(row.get("칼라명",     "") or "").strip()
            size    = str(row.get("사이즈명",   "") or "").strip()
            if not barcode and not name:
                continue
            raw_products_off.append({
                "date": date, "barcode": barcode, "name": name,
                "color": color, "size": size,
                "price":  _int(row.get("판매단가",  0)),
                "qty":    _int(row.get("판매수량",  0)),
                "amount": _int(row.get("실판매금액", 0)),
            })
        print(f"상품별매출(off) {len(raw_products_off)}행 로드")
    except Exception as e:
        print(f"상품별매출(off) 로드 오류: {e}")

    # ── 상품별매출(on) ────────────────────────────────────
    raw_products_on = []
    try:
        ws_on  = sh.worksheet("상품별매출(on)")
        on_raw = ws_on.get_all_records()
        for row in on_raw:
            date    = _norm_date(row.get("판매일자", "") or "")
            barcode = _barcode(row.get("바코드",""))
            name    = str(row.get("상품명",   "") or "").strip()
            size    = str(row.get("사이즈",   "") or "").strip()
            player  = str(row.get("선수명",   "") or "").strip()
            if not barcode and not name:
                continue
            raw_products_on.append({
                "date": date, "barcode": barcode, "name": name,
                "size": size, "player": player,
                "price":  _int(row.get("판매단가",  0)),
                "qty":    _int(row.get("판매수량",  0)),
                "amount": _int(row.get("실판매금액", 0)),
            })
        print(f"상품별매출(on) {len(raw_products_on)}행 로드")
    except Exception as e:
        print(f"상품별매출(on) 로드 오류: {e}")

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

    return merged, news, digest, raw_products_off, raw_products_on


def build_html(data: list, news: list, digest: str, raw_products_off: list, raw_products_on: list, apps_script_url: str = "") -> str:
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

    # 월별 홈/어웨이·결과별 합산
    from collections import defaultdict
    def month_label(d): return d[:4] + "." + d[5:7]  # "2026.04"

    monthly_ha_sum  = defaultdict(lambda: {"홈": {"off": 0, "on": 0}, "어웨이": {"off": 0, "on": 0}})
    monthly_res_sum = defaultdict(lambda: {"승": {"off": 0, "on": 0}, "패": {"off": 0, "on": 0}})
    for r in game_days:
        m  = month_label(r["date"])
        ha = r["home_away"]
        rs = r["result"]
        if ha in ["홈", "어웨이"]:
            monthly_ha_sum[m][ha]["off"] += r["off"]
            monthly_ha_sum[m][ha]["on"]  += r["on"]
        if rs in ["승", "패"]:
            monthly_res_sum[m][rs]["off"] += r["off"]
            monthly_res_sum[m][rs]["on"]  += r["on"]

    ha_months  = sorted(monthly_ha_sum.keys())
    res_months = sorted(monthly_res_sum.keys())

    ha_month_labels  = [m[5:].lstrip("0") + "월" for m in ha_months]
    res_month_labels = [m[5:].lstrip("0") + "월" for m in res_months]

    home_total = [monthly_ha_sum[m]["홈"]["off"]    + monthly_ha_sum[m]["홈"]["on"]    for m in ha_months]
    away_total = [monthly_ha_sum[m]["어웨이"]["off"] + monthly_ha_sum[m]["어웨이"]["on"] for m in ha_months]
    win_total  = [monthly_res_sum[m]["승"]["off"] + monthly_res_sum[m]["승"]["on"] for m in res_months]
    lose_total = [monthly_res_sum[m]["패"]["off"] + monthly_res_sum[m]["패"]["on"] for m in res_months]

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

    raw_off_json = json.dumps(raw_products_off, ensure_ascii=False)
    raw_on_json  = json.dumps(raw_products_on,  ensure_ascii=False)

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
    flex: 1; min-width: 140px; background: white;
    border-radius: 12px; padding: 20px 24px;
    box-shadow: 0 2px 8px rgba(0,0,0,.07);
  }}
  .kpi.wide {{ flex: 2; min-width: 200px; }}
  .kpi.narrow {{ flex: 0.7; min-width: 110px; }}
  .kpi .label {{ font-size: 12px; color: #888; margin-bottom: 8px; }}
  .kpi .value {{
    font-size: 22px; font-weight: 700; color: #002D72;
    display: flex; align-items: baseline; gap: 3px; white-space: nowrap;
  }}
  .kpi .value.red {{ color: #C8102E; }}
  .kpi .value .unit {{ font-size: 13px; font-weight: 400; color: #888; }}
  .kpi .sub   {{ font-size: 11px; color: #aaa; margin-top: 4px; }}

  /* 차트 */
  .charts-top {{ padding: 20px 32px 0; }}
  .charts-bottom {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; padding: 16px 32px 0; }}
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
  .drilldown-overlay {{ display:none; position:fixed; inset:0; background:rgba(0,0,0,0.45); z-index:1000; align-items:center; justify-content:center; }}
  .drilldown-overlay.open {{ display:flex; }}
  .drilldown-modal {{ background:#fff; border-radius:14px; padding:28px; width:95%; max-width:1200px; max-height:85vh; overflow-y:auto; box-shadow:0 8px 40px rgba(0,0,0,0.2); position:relative; }}
  .drilldown-charts {{ display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-top:16px; }}
  .drilldown-chart-wrap {{ height:220px; position:relative; }}

  @media (max-width: 900px) {{
    .charts-bottom {{ grid-template-columns: 1fr; }}
    .kpi-row, .charts-top, .charts-bottom, .table-section, .product-section, .issue-section {{ padding-left: 16px; padding-right: 16px; }}
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
  <div class="kpi wide">
    <div class="label">OFF 거래액 (누계)</div>
    <div class="value"><span id="kpi-off">{fmt(total_off)}</span><span class="unit">원</span></div>
    <div class="sub">{date_range_label}</div>
  </div>
  <div class="kpi wide">
    <div class="label">ON 거래액 (누계)</div>
    <div class="value"><span id="kpi-on">{fmt(total_on)}</span><span class="unit">원</span></div>
    <div class="sub">{date_range_label}</div>
  </div>
  <div class="kpi wide">
    <div class="label">총 거래액 (누계)</div>
    <div class="value red"><span id="kpi-total">{fmt(total_total)}</span><span class="unit">원</span></div>
    <div class="sub">{date_range_label}</div>
  </div>
  <div class="kpi narrow">
    <div class="label">경기 수</div>
    <div class="value">{len(game_days)}</div>
    <div class="sub">홈 {len(home_rows)} · 어웨이 {len(away_rows)}</div>
  </div>
  <div class="kpi narrow">
    <div class="label">승률</div>
    <div class="value">{win_rate}</div>
    <div class="sub">{wins}승 {losses}패</div>
  </div>
  <div class="kpi narrow">
    <div class="label">홈 평균 관중</div>
    <div class="value">{fmt(avg_crowd)}<span class="unit">명</span></div>
    <div class="sub">점유율 {avg_occupancy}</div>
  </div>
</div>

<div class="charts-top">
  <div class="chart-card" style="height:320px">
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
</div>
<div class="charts-bottom">
  <div class="chart-card" style="height:280px">
    <h3>홈 / 어웨이 월별 거래액</h3>
    <div style="position:relative;width:100%;height:calc(100% - 36px)">
      <canvas id="haChart"></canvas>
    </div>
  </div>
  <div class="chart-card" style="height:280px">
    <h3>경기 결과별 월별 거래액</h3>
    <div style="position:relative;width:100%;height:calc(100% - 36px)">
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
        <h3 style="margin-bottom:4px">상품별 누적 판매 실적 (온/오프 통합)</h3>
        <span id="productRangeLabel" style="font-size:11px;color:#aaa"></span>
        <span style="font-size:11px;color:#bbb;margin-left:8px">※ 온라인 판매수량·금액은 취소 미반영</span>
      </div>
      <div style="display:flex;gap:8px">
        <button id="productExcelBtn" onclick="downloadProductExcel()">📥 엑셀 다운로드</button>
        <button onclick="window.open('https://melonredash.melon.com/queries/16900/source?p_%EC%A1%B0%ED%9A%8C%20%EA%B8%B0%EA%B0%84=2026-02-24--2026-06-16','_blank')" style="background:#f0f0f0;border:none;border-radius:8px;padding:6px 14px;cursor:pointer;font-size:12px;font-weight:600;">📊 온라인 실적 조회</button>
        {'<button onclick="openUploadModal()" style="background:#C8102E;color:#fff;border:none;border-radius:8px;padding:6px 14px;cursor:pointer;font-size:12px;font-weight:600;">📤 온라인 실적 업로드</button>' if apps_script_url else ''}
      </div>
    </div>
    <div class="product-controls">
      <input type="date" id="prodStartDate" title="시작일">
      <span style="font-size:12px;color:#888">~</span>
      <input type="date" id="prodEndDate" title="종료일">
      <input type="text" id="prodNameSearch" placeholder="상품명 검색" style="width:140px">
      <input type="text" id="prodCodeSearch" placeholder="바코드 (,로 여러개)" style="width:180px">
      <button class="ctrl-btn ctrl-btn-primary" onclick="applyProductFilter()">조회</button>
      <button class="ctrl-btn ctrl-btn-reset" onclick="resetProductFilter()">초기화</button>
    </div>
    <p class="product-guide">💡 기간·상품명·바코드 필터 후 조회하거나, 엑셀 다운로드로 날짜별 전체 데이터를 확인하세요.</p>
    <table id="productTable" style="font-size:12px">
      <thead>
        <tr>
          <th>바코드</th>
          <th>상품명</th>
          <th>칼라</th><th>사이즈</th><th>선수명</th>
          <th style="text-align:right">판매단가</th>
          <th style="text-align:right">OFF수량</th>
          <th style="text-align:right">ON수량</th>
          <th style="text-align:right">합계수량</th>
          <th style="text-align:right">OFF금액</th>
          <th style="text-align:right">ON금액</th>
          <th style="text-align:right">합계금액</th>
          <th style="text-align:center">트렌드</th>
        </tr>
      </thead>
      <tbody id="productTbody"></tbody>
    </table>
    <div class="pagination" id="productPagination"></div>
  </div>
</div>

<div class="drilldown-overlay" id="drilldownSection" onclick="handleOverlayClick(event)">
  <div class="drilldown-modal">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
      <div>
        <h3 id="drilldownTitle" style="margin-bottom:4px"></h3>
        <span id="drilldownSub" style="font-size:11px;color:#aaa"></span>
      </div>
      <button onclick="closeDrilldown()" style="background:#f0f0f0;border:none;border-radius:8px;padding:6px 14px;cursor:pointer;font-size:12px;font-weight:600;">✕ 닫기</button>
    </div>
    <div id="drilldownStats" style="display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap"></div>
    <div class="drilldown-charts">
      <div>
        <div style="font-size:12px;color:#555;font-weight:600;margin-bottom:8px">일별 판매수량</div>
        <div class="drilldown-chart-wrap"><canvas id="drilldownQtyChart"></canvas></div>
      </div>
      <div>
        <div style="font-size:12px;color:#555;font-weight:600;margin-bottom:8px">일별 실판매금액</div>
        <div class="drilldown-chart-wrap"><canvas id="drilldownAmountChart"></canvas></div>
      </div>
    </div>
  </div>
</div>

<!-- 업로드 모달 -->
<div class="drilldown-overlay" id="uploadModal" onclick="handleUploadOverlayClick(event)" {'style="display:none"' if not apps_script_url else ''}>
  <div class="drilldown-modal" style="max-width:480px">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px">
      <h3 style="margin:0">온라인 실적 업로드</h3>
      <button onclick="closeUploadModal()" style="background:#f0f0f0;border:none;border-radius:8px;padding:6px 14px;cursor:pointer;font-size:12px;font-weight:600;">✕ 닫기</button>
    </div>
    <div style="margin-bottom:14px">
      <label style="font-size:12px;color:#555;font-weight:600;display:block;margin-bottom:6px">엑셀 파일 선택</label>
      <input type="file" id="uploadFileInput" accept=".xlsx,.xls,.csv" style="font-size:13px;width:100%">
    </div>
    <div style="margin-bottom:20px">
      <label style="font-size:12px;color:#555;font-weight:600;display:block;margin-bottom:6px">비밀번호</label>
      <input type="password" id="uploadKeyInput" placeholder="업로드 비밀번호 입력" style="width:100%;padding:8px 10px;border:1px solid #ddd;border-radius:8px;font-size:13px;box-sizing:border-box">
    </div>
    <div id="uploadStatus" style="font-size:12px;color:#888;margin-bottom:14px;min-height:18px"></div>
    <button onclick="submitUpload()" style="width:100%;background:#002D72;color:#fff;border:none;border-radius:8px;padding:10px;cursor:pointer;font-size:14px;font-weight:700;">업로드</button>
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

// ── 상품별 매출 (온/오프 통합) ──
const rawOffData = {raw_off_json};
const rawOnData  = {raw_on_json};
const PRODUCT_ROWS_PER_PAGE = 10;
let productPage = 1;
let currentProductRows = [];
let drilldownAmtChart = null;
let drilldownQtyChart = null;

function dotToHyphen(d) {{ return d ? d.replace(/\./g, '-') : ''; }}
function hyphenToDot(d) {{ return d ? d.replace(/-/g, '.') : ''; }}

function getFilters() {{
  return {{
    start:    document.getElementById('prodStartDate').value,
    end:      document.getElementById('prodEndDate').value,
    name:     document.getElementById('prodNameSearch').value.trim(),
    barcodes: document.getElementById('prodCodeSearch').value
                .split(',').map(s => s.trim()).filter(s => s),
  }};
}}

function filterRows(rows, f, nameKey='name', barcodeKey='barcode') {{
  return rows.filter(r => {{
    if (f.start && dotToHyphen(r.date) < f.start) return false;
    if (f.end   && dotToHyphen(r.date) > f.end)   return false;
    if (f.name  && !(r[nameKey]||'').includes(f.name)) return false;
    if (f.barcodes.length && !f.barcodes.some(c => (r[barcodeKey]||'').includes(c))) return false;
    return true;
  }});
}}

const PLAYER_KEYWORDS = ['구자욱','오승환','최형우','박승규','디아즈','김영웅','원태인','강민호',
  '류지혁','이재현','전준우','김지찬','김성윤','김재윤','김태훈','김헌곤','매닝','미야지',
  '박세혁','배찬승','백정현','심재훈','양창섭','이성규','이승현','이호성','임기영','장찬희',
  '최원태','최지광','함수호','후라도'];

function extractPlayerFromName(name) {{
  if (!name) return '';
  for (const p of PLAYER_KEYWORDS) {{
    if (name.includes(p)) return p;
  }}
  return '';
}}

const COLOR_KEYWORDS = ['NAVY','BLUE','WHITE','BLACK','RED','PINK','GRAY','GREY','GREEN',
  'YELLOW','ORANGE','PURPLE','BROWN','BEIGE','IVORY','KHAKI','MINT','WINE',
  'SKY BLUE','MELANGE GREY','LIGHT GREY','CHARCOAL','CREAM','GOLD','SILVER',
  '블루','화이트','블랙','네이비','레드','핑크','그레이','그린','민트','베이지','아이보리'];

function extractColorFromName(name) {{
  if (!name) return '';
  // 괄호 안 색상 추출: (NAVY), (SKY BLUE) 등
  const m = name.match(/\(([A-Z가-힣][A-Z가-힣 ]*)\)/);
  if (m) {{
    const candidate = m[1].trim();
    // 색상 키워드 포함 여부 확인 (사람이름·이벤트문구 제외)
    if (COLOR_KEYWORDS.some(c => candidate.toUpperCase().includes(c.toUpperCase()))) {{
      return candidate;
    }}
  }}
  // 상품명 내 색상 키워드 직접 추출
  for (const c of COLOR_KEYWORDS) {{
    const re = new RegExp('\\\\b' + c + '\\\\b', 'i');
    if (re.test(name)) return c.toUpperCase();
  }}
  return '';
}}

const isFree = v => !v || v==='-' || v.trim().toLowerCase()==='free' || v.trim()==='공통';

function mergeProducts(offRows, onRows) {{
  const map = {{}};

  offRows.forEach(r => {{
    if (!r.barcode) return;
    const k = r.barcode;
    if (!map[k]) map[k] = {{
      barcode: k, off_name: r.name||'-', on_name: '-',
      color: r.color||'-', size: r.size||'-', player: extractPlayerFromName(r.name)||'-',
      price: r.price, off_qty: 0, off_amount: 0, on_qty: 0, on_amount: 0,
    }};
    map[k].off_qty    += r.qty;
    map[k].off_amount += r.amount;
  }});
  onRows.forEach(r => {{
    if (!r.barcode) return;
    const k = r.barcode;
    if (!map[k]) {{
      const nameColor  = extractColorFromName(r.name);
      const namePlayer = isFree(r.player) ? extractPlayerFromName(r.name) : r.player;
      map[k] = {{
        barcode: k, off_name: '-', on_name: r.name||'-',
        color: nameColor||'-', size: r.size||'-', player: namePlayer||'-',
        price: r.price, off_qty: 0, off_amount: 0, on_qty: 0, on_amount: 0,
      }};
    }} else {{
      if (r.name) map[k].on_name = r.name;
      // OFF 사이즈가 free/공통이고 ON에 실제 사이즈 있으면 ON 우선
      if (isFree(map[k].size) && !isFree(r.size)) map[k].size = r.size;
      // OFF 색상이 free/공통이고 ON에 실제 색상 있으면 ON 우선
      const onColor = isFree(r.color) ? extractColorFromName(r.name) : r.color;
      if (isFree(map[k].color) && onColor) map[k].color = onColor;
      // OFF 선수명 없고 ON에 선수명 있으면 ON 우선 (필드 또는 상품명 추출)
      const onPlayer = isFree(r.player) ? extractPlayerFromName(r.name) : r.player;
      if (isFree(map[k].player) && onPlayer) map[k].player = onPlayer;
      if (!map[k].price) map[k].price = r.price;
    }}
    map[k].on_qty    += r.qty;
    map[k].on_amount += r.amount;
  }});
  return Object.values(map)
    .sort((a,b) => (b.off_amount+b.on_amount) - (a.off_amount+a.on_amount));
}}

function applyProductFilter() {{
  const f = getFilters();
  const filteredOff = filterRows(rawOffData, f);
  const filteredOn  = filterRows(rawOnData,  f);
  currentProductRows = mergeProducts(filteredOff, filteredOn);
  productPage = 1;
  let label = '';
  if (f.start||f.end) label += (hyphenToDot(f.start)||'전체') + ' ~ ' + (hyphenToDot(f.end)||'전체');
  if (f.name)         label += (label?' / ':'') + '상품명: ' + f.name;
  if (f.barcodes.length) label += (label?' / ':'') + '바코드: ' + f.barcodes.join(', ');
  document.getElementById('productRangeLabel').textContent = label || '';
  renderProductTable();
}}

function resetProductFilter() {{
  document.getElementById('prodStartDate').value  = '';
  document.getElementById('prodEndDate').value    = '';
  document.getElementById('prodNameSearch').value = '';
  document.getElementById('prodCodeSearch').value = '';
  currentProductRows = mergeProducts(rawOffData, rawOnData);
  productPage = 1;
  document.getElementById('productRangeLabel').textContent = '전체 기간 합산';
  renderProductTable();
}}

function renderProductTable() {{
  const tbody      = document.getElementById('productTbody');
  const total      = currentProductRows.length;
  const totalPages = Math.ceil(total / PRODUCT_ROWS_PER_PAGE);
  const start      = (productPage - 1) * PRODUCT_ROWS_PER_PAGE;
  const pageRows   = currentProductRows.slice(start, start + PRODUCT_ROWS_PER_PAGE);
  const fmt = n => (n && n!==0) ? n.toLocaleString('ko-KR') : '-';

  const totOffQty = currentProductRows.reduce((s,p)=>s+p.off_qty,0);
  const totOffAmt = currentProductRows.reduce((s,p)=>s+p.off_amount,0);
  const totOnQty  = currentProductRows.reduce((s,p)=>s+p.on_qty,0);
  const totOnAmt  = currentProductRows.reduce((s,p)=>s+p.on_amount,0);

  let html = '';
  pageRows.forEach((p, idx) => {{
    const total_amt = p.off_amount + p.on_amount;
    const opts = [p.color,p.size,p.player].filter(v=>v&&v!=='-').join(' / ');
    window.__ddOpts = window.__ddOpts || {{}};
    window.__ddOpts[p.barcode] = opts;
    html += `<tr style="cursor:pointer" onclick="showDrilldown('${{p.barcode.replace(/'/g,"\\'")}}','${{(p.off_name||p.on_name).replace(/'/g,"\\'")}}')">
      <td style="font-size:11px;color:#888">${{p.barcode}}</td>
      <td><span style="font-size:10px;color:#888;display:inline-block;width:30px">[OFF]</span><span style="font-size:11px">${{p.off_name}}</span><br><span style="font-size:10px;color:#888;display:inline-block;width:30px">[ON]</span><span style="font-size:11px">${{p.on_name}}</span></td>
      <td>${{isFree(p.color)?'-':p.color}}</td><td>${{isFree(p.size)?'-':p.size}}</td><td>${{isFree(p.player)?'-':p.player}}</td>
      <td class="num">${{fmt(p.price)}}</td>
      <td class="num">${{fmt(p.off_qty)}}</td>
      <td class="num">${{fmt(p.on_qty)}}</td>
      <td class="num bold">${{fmt(p.off_qty + p.on_qty)}}</td>
      <td class="num">${{fmt(p.off_amount)}}</td>
      <td class="num">${{fmt(p.on_amount)}}</td>
      <td class="num bold">${{fmt(total_amt)}}</td>
      <td style="text-align:center;color:#002D72">📈</td>
    </tr>`;
  }});
  html += `<tr class="product-total-row">
    <td colspan="6">합계 (전체 ${{total}}개 SKU)</td>
    <td class="num">${{fmt(totOffQty)}}</td>
    <td class="num">${{fmt(totOnQty)}}</td>
    <td class="num bold">${{fmt(totOffQty+totOnQty)}}</td>
    <td class="num">${{fmt(totOffAmt)}}</td>
    <td class="num">${{fmt(totOnAmt)}}</td>
    <td class="num bold">${{fmt(totOffAmt+totOnAmt)}}</td>
    <td></td>
  </tr>`;
  tbody.innerHTML = html;

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

// ── 드릴다운: OFF/ON 별도 라인 ──
function showDrilldown(barcode, label) {{
  const opts = (window.__ddOpts && window.__ddOpts[barcode]) || '';
  const f = getFilters();
  const offRows = filterRows(rawOffData, f).filter(r => r.barcode === barcode);
  const onRows  = filterRows(rawOnData,  f).filter(r => r.barcode === barcode);

  const offByDate = {{}}, onByDate = {{}};
  offRows.forEach(r => {{
    if (!offByDate[r.date]) offByDate[r.date] = {{qty:0, amount:0}};
    offByDate[r.date].qty    += r.qty;
    offByDate[r.date].amount += r.amount;
  }});
  onRows.forEach(r => {{
    if (!onByDate[r.date]) onByDate[r.date] = {{qty:0, amount:0}};
    onByDate[r.date].qty    += r.qty;
    onByDate[r.date].amount += r.amount;
  }});

  const allDates = [...new Set([...Object.keys(offByDate), ...Object.keys(onByDate)])].sort();
  const labels   = allDates.map(d => d.slice(5));
  const offAmts  = allDates.map(d => offByDate[d] ? offByDate[d].amount : null);
  const onAmts   = allDates.map(d => onByDate[d]  ? onByDate[d].amount  : null);
  const offQtys  = allDates.map(d => offByDate[d] ? offByDate[d].qty    : null);
  const onQtys   = allDates.map(d => onByDate[d]  ? onByDate[d].qty     : null);

  const totOffQty = offQtys.reduce((s,v)=>s+v,0);
  const totOnQty  = onQtys.reduce((s,v)=>s+v,0);
  const totOffAmt = offAmts.reduce((s,v)=>s+v,0);
  const totOnAmt  = onAmts.reduce((s,v)=>s+v,0);
  const fmt = n => n.toLocaleString('ko-KR');
  const pct = (n, total) => total > 0 ? (n/total*100).toFixed(1)+'%' : '-';
  const statCard = (label, val, color, sub='') =>
    `<div style="background:#f8f9fa;border-radius:10px;padding:10px 18px;min-width:130px">
      <div style="font-size:11px;color:#888;margin-bottom:3px">${{label}}</div>
      <div style="font-size:16px;font-weight:700;color:${{color}}">${{val}}</div>
      ${{sub ? `<div style="font-size:11px;color:#aaa;margin-top:2px">${{sub}}</div>` : ''}}
    </div>`;
  const totQty = totOffQty + totOnQty;
  const totAmt = totOffAmt + totOnAmt;
  document.getElementById('drilldownStats').innerHTML =
    statCard('OFF 판매수량', fmt(totOffQty)+'개', '#002D72', '구성비 '+pct(totOffQty, totQty)) +
    statCard('ON 판매수량',  fmt(totOnQty)+'개',  '#C8102E', '구성비 '+pct(totOnQty,  totQty)) +
    statCard('합계수량', fmt(totQty)+'개', '#333') +
    statCard('OFF 판매금액', fmt(totOffAmt)+'원', '#002D72', '구성비 '+pct(totOffAmt, totAmt)) +
    statCard('ON 판매금액',  fmt(totOnAmt)+'원',  '#C8102E', '구성비 '+pct(totOnAmt,  totAmt)) +
    statCard('합계금액', fmt(totAmt)+'원', '#333') +
    ``;

  document.getElementById('drilldownTitle').innerHTML =
    `${{label}}<span style="font-size:13px;font-weight:400;color:#666;margin-left:8px">${{opts}}</span>`;
  document.getElementById('drilldownSub').textContent   = `${{barcode}} · ${{allDates.length}}일 판매`;
  document.getElementById('drilldownSection').classList.add('open');

  if (drilldownAmtChart) drilldownAmtChart.destroy();
  if (drilldownQtyChart) drilldownQtyChart.destroy();

  const amtOpts = {{ responsive:true, maintainAspectRatio:false, spanGaps:true,
    plugins:{{ legend:{{ position:'bottom' }} }},
    scales:{{ y:{{ ticks:{{ callback: v => v>=1e6?(v/1e6).toFixed(1)+'M':v.toLocaleString() }} }} }} }};

  drilldownAmtChart = new Chart(document.getElementById('drilldownAmountChart'), {{
    type: 'line',
    data: {{ labels, datasets: [
      {{ label:'OFF 실판매금액', data:offAmts,
         borderColor:OFF, backgroundColor:'rgba(0,45,114,0.08)',
         fill:false, tension:0.3, pointRadius:3, borderWidth:2 }},
      {{ label:'ON 실판매금액',  data:onAmts,
         borderColor:ON,  backgroundColor:'rgba(200,16,46,0.08)',
         fill:false, tension:0.3, pointRadius:3, borderWidth:2, }},
    ]}},
    options: amtOpts,
  }});

  drilldownQtyChart = new Chart(document.getElementById('drilldownQtyChart'), {{
    type: 'line',
    data: {{ labels, datasets: [
      {{ label:'OFF 판매수량', data:offQtys,
         borderColor:OFF, backgroundColor:'rgba(0,45,114,0.08)',
         fill:false, tension:0.3, pointRadius:3, borderWidth:2 }},
      {{ label:'ON 판매수량',  data:onQtys,
         borderColor:ON,  backgroundColor:'rgba(200,16,46,0.08)',
         fill:false, tension:0.3, pointRadius:3, borderWidth:2, }},
    ]}},
    options: {{ responsive:true, maintainAspectRatio:false, spanGaps:true,
      plugins:{{ legend:{{ position:'bottom' }} }},
      scales:{{ y:{{ ticks:{{ callback: v => v.toLocaleString() }} }} }} }},
  }});
}}

function closeDrilldown() {{
  document.getElementById('drilldownSection').classList.remove('open');
  if (drilldownAmtChart) {{ drilldownAmtChart.destroy(); drilldownAmtChart = null; }}
  if (drilldownQtyChart) {{ drilldownQtyChart.destroy(); drilldownQtyChart = null; }}
}}

function handleOverlayClick(e) {{
  if (e.target === document.getElementById('drilldownSection')) closeDrilldown();
}}

// ── 온라인 실적 업로드 ──
const APPS_SCRIPT_URL = '{apps_script_url}';

function openUploadModal() {{
  document.getElementById('uploadModal').classList.add('open');
  document.getElementById('uploadStatus').textContent = '';
  document.getElementById('uploadFileInput').value = '';
}}
function closeUploadModal() {{
  document.getElementById('uploadModal').classList.remove('open');
}}
function handleUploadOverlayClick(e) {{
  if (e.target === document.getElementById('uploadModal')) closeUploadModal();
}}

async function submitUpload() {{
  const fileInput = document.getElementById('uploadFileInput');
  const key = document.getElementById('uploadKeyInput').value.trim();
  const status = document.getElementById('uploadStatus');

  if (!fileInput.files.length) {{ status.textContent = '❌ 파일을 선택해주세요.'; return; }}
  if (!key) {{ status.textContent = '❌ 비밀번호를 입력해주세요.'; return; }}

  status.textContent = '⏳ 파일 파싱 중...';
  const file = fileInput.files[0];
  const ab = await file.arrayBuffer();
  const wb = XLSX.read(ab, {{ type: 'array' }});
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, {{ defval: '' }});

  if (!raw.length) {{ status.textContent = '❌ 데이터가 없습니다.'; return; }}

  // 컬럼 매핑 (CSV헤더 → 시트헤더)
  const COL_MAP = {{
    '판매일':   '판매일자', '판매일자': '판매일자',
    '상품ID':   '상품ID',
    '상품명':   '상품명',
    '바코드':   '바코드',
    'skucode':  'skucode', 'SKUcode': 'skucode',
    '사이즈':   '사이즈',
    '선수명':   '선수명',
    '판매가':   '판매단가', '판매단가': '판매단가',
    '결제상품수': '판매수량', '판매수량': '판매수량',
    '상품결제금액': '실판매금액', '실판매금액': '실판매금액',
  }};
  const HEADER = ['판매일자','상품ID','상품명','바코드','skucode','사이즈','선수명','판매단가','판매수량','실판매금액'];

  const rows = raw.map(r => {{
    const mapped = {{}};
    Object.entries(r).forEach(([k, v]) => {{
      const target = COL_MAP[k.trim()];
      if (target) mapped[target] = String(v).trim();
    }});
    return mapped;
  }}).filter(r => r['판매일자']);

  if (!rows.length) {{ status.textContent = '❌ 판매일자 컬럼을 찾을 수 없습니다.'; return; }}

  const dates = [...new Set(rows.map(r => r['판매일자']))];
  status.textContent = `⏳ ${{rows.length}}행 (${{dates.length}}일) 업로드 중...`;

  try {{
    const resp = await fetch(APPS_SCRIPT_URL, {{
      method: 'POST',
      body: JSON.stringify({{ key, rows }}),
    }});
    const result = await resp.json();
    if (result.ok) {{
      status.style.color = '#1a7f37';
      status.textContent = `✅ ${{result.inserted}}행 업로드 완료! 대시보드 재생성 중...`;
    }} else {{
      status.style.color = '#C8102E';
      status.textContent = '❌ ' + result.error;
    }}
  }} catch(err) {{
    status.style.color = '#C8102E';
    status.textContent = '❌ 업로드 실패: ' + err.message;
  }}
}}

// ── 엑셀 다운로드 (OFF + ON 날짜별 전체) ──
function downloadProductExcel() {{
  const f = getFilters();
  const offRows = filterRows(rawOffData, f);
  const onRows  = filterRows(rawOnData,  f);

  // mergeProducts로 바코드별 옵션값 확보
  const merged = {{}};
  mergeProducts(offRows, onRows).forEach(p => {{ merged[p.barcode] = p; }});

  const offSheet = [['판매일자','바코드','OFF상품명','칼라','사이즈','판매단가','판매수량','실판매금액'],
    ...offRows.sort((a,b)=>a.date.localeCompare(b.date))
      .map(r => {{
        const m = merged[r.barcode] || {{}};
        return [r.date, r.barcode, r.name,
          m.color||r.color, m.size||r.size,
          r.price, r.qty, r.amount];
      }})];

  const onSheet = [['판매일자','바코드','ON상품명','칼라','사이즈','선수명','판매단가','판매수량','실판매금액'],
    ...onRows.sort((a,b)=>a.date.localeCompare(b.date))
      .map(r => {{
        const m = merged[r.barcode] || {{}};
        return [r.date, r.barcode, r.name,
          m.color || extractColorFromName(r.name),
          m.size  || r.size,
          m.player|| extractPlayerFromName(r.name) || r.player,
          r.price, r.qty, r.amount];
      }})];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(offSheet), '오프라인');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(onSheet),  '온라인');
  XLSX.writeFile(wb, '삼성라이온즈_상품별매출_온오프.xlsx');
}}

// 초기 렌더링: 전체
currentProductRows = mergeProducts(rawOffData, rawOnData);
document.getElementById('productRangeLabel').textContent = '전체 기간 합산';
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
        ctx.fillText(Math.round(v / 1e6).toLocaleString('ko-KR') + '백만', bar.x, bar.y - 5);
        ctx.restore();
      }});
    }});
  }}
}};

const barSideOpts = () => ({{
  responsive: true,
  maintainAspectRatio: false,
  plugins: {{ legend: {{ position: 'bottom', labels: {{ boxWidth: 12, font: {{ size: 11 }} }} }} }},
  scales: {{
    y: {{
      min: 0,
      ticks: {{ callback: v => v >= 1e6 ? (v/1e6).toFixed(0)+'M' : v.toLocaleString(), font: {{ size: 10 }} }},
      grid: {{ color: '#f0f0f0' }},
    }},
    x: {{ ticks: {{ font: {{ size: 12 }} }} }}
  }},
  layout: {{ padding: {{ top: 16 }} }}
}});

// ── 홈/어웨이 월별 합산 막대 ──
new Chart(document.getElementById('haChart'), {{
  type: 'bar',
  plugins: [topLabelPlugin],
  data: {{
    labels: {json.dumps(ha_month_labels)},
    datasets: [
      {{ label: '홈',    data: {json.dumps(home_total)}, backgroundColor: alpha(OFF, .85) }},
      {{ label: '어웨이', data: {json.dumps(away_total)}, backgroundColor: alpha(ON,  .75) }},
    ]
  }},
  options: barSideOpts(),
}});

// ── 결과별 월별 합산 막대 ──
new Chart(document.getElementById('resultChart'), {{
  type: 'bar',
  plugins: [topLabelPlugin],
  data: {{
    labels: {json.dumps(res_month_labels)},
    datasets: [
      {{ label: '승', data: {json.dumps(win_total)},  backgroundColor: 'rgba(46,125,50,.8)' }},
      {{ label: '패', data: {json.dumps(lose_total)}, backgroundColor: 'rgba(198,40,40,.7)' }},
    ]
  }},
  options: barSideOpts(),
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
    data, news, digest, raw_products_off, raw_products_on = fetch_data()
    print(f"병합된 데이터: {len(data)}일 / 뉴스이슈: {len(news)}건 / OFF: {len(raw_products_off)}행 / ON: {len(raw_products_on)}행")

    os.makedirs("dashboard", exist_ok=True)
    apps_script_url = os.environ.get("APPS_SCRIPT_URL", "https://script.google.com/macros/s/AKfycbxAmSM8usUuNTXAEUdOf1Z3uCmIlEc1nADzB9IHxbK8pmf3mFncePDikk4AXN46Ygc-Hw/exec")
    html = build_html(data, news, digest, raw_products_off, raw_products_on, apps_script_url)
    with open("dashboard/index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("dashboard/index.html 생성 완료")


if __name__ == "__main__":
    main()
