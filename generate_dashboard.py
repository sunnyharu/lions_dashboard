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
        off = row.get("OFF거래액", "") or 0
        on  = row.get("ON거래액",  "") or 0
        try: off = int(float(str(off).replace(",", "")))
        except: off = 0
        try: on  = int(float(str(on).replace(",", "")))
        except: on  = 0
        # 날짜 중복 시 ON거래액이 있는 행 우선
        if date in sales and on == 0:
            continue
        sales[date] = {"off": off, "on": on}

    # ── 뉴스이슈 ──────────────────────────────────────────
    news = []
    try:
        ws_news  = sh.worksheet("뉴스이슈")
        news_raw = ws_news.get_all_records()
        for row in news_raw:
            date = str(row.get("날짜", "")).strip()
            if not date:
                continue
            parts = date.split(".")
            if len(parts) == 3:
                date = f"{parts[0]}.{int(parts[1]):02d}.{int(parts[2]):02d}"
            news.append({
                "date":    date,
                "source":  row.get("출처", ""),
                "title":   row.get("제목", ""),
                "summary": row.get("AI요약", "") or row.get("요약", ""),
                "views":   row.get("조회수", 0) or 0,
                "link":    row.get("링크", ""),
            })
    except Exception:
        pass

    # ── 카페트렌드 다이제스트 ─────────────────────────────
    digest = ""
    try:
        ws_digest  = sh.worksheet("카페트렌드")
        digest_raw = ws_digest.get_all_values()
        if len(digest_raw) > 1:
            digest = digest_raw[-1][1] if len(digest_raw[-1]) > 1 else ""
    except Exception:
        pass

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
        games[date] = {
            "home_away": row.get("홈/어웨이", ""),
            "opponent":  row.get("상대팀", ""),
            "result":    row.get("결과", ""),
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
        })

    return merged, news, digest


def build_html(data: list, news: list, digest: str) -> str:
    game_days = [r for r in data if r["result"]]

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

    # Chart 데이터
    chart_dates  = [r["date"][5:] for r in data]  # "04.01" 형태
    chart_off    = [r["off"]   for r in data]
    chart_on     = [r["on"]    for r in data]
    chart_total  = [r["total"] for r in data]

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

    # 뉴스 패널 HTML
    from collections import defaultdict as _dd
    news_by_date = _dd(list)
    for n in news:
        news_by_date[n["date"]].append(n)

    sorted_dates = sorted(news_by_date.keys(), reverse=True)

    # 날짜 필터 버튼
    filter_buttons = '<button class="active" onclick="filterNews(\'all\', this)">전체</button>'
    for d in sorted_dates:
        label = d[5:]  # "04.28"
        filter_buttons += f'<button onclick="filterNews(\'{d}\', this)">{label}</button>'

    # 날짜별 뉴스 블록
    news_blocks = ""
    for date in sorted_dates:
        items = news_by_date[date]
        news_blocks += f'<div class="news-group" data-date="{date}">'
        news_blocks += f'<div class="news-date">{date}</div>'
        for n in items:
            src_cls    = "badge-news" if n["source"] == "뉴스" else "badge-cafe"
            link_open  = f'<a href="{n["link"]}" target="_blank" rel="noopener">' if n["link"] else ""
            link_close = "</a>" if n["link"] else ""
            views_html = f'<span class="news-views">조회 {int(n.get("views",0)):,}</span>' if n.get("views") else ""
        news_blocks += f"""
            <div class="news-item">
              <span class="news-src {src_cls}">{n["source"]}</span>{views_html}
              {link_open}<span class="news-title">{n["title"]}</span>{link_close}
              {'<div class="news-desc">' + n["summary"] + '</div>' if n["summary"] else ""}
            </div>"""
        news_blocks += '</div>'

    if not news_blocks:
        news_blocks = '<div class="news-empty">수집된 이슈 없음</div>'
        filter_buttons = ""

    digest_html = ""
    if digest:
        digest_html = f'''<div class="digest-box">
          <div class="digest-label">최근 7일 카페 트렌드 다이제스트</div>
          <div class="digest-text">{digest}</div>
        </div>'''

    news_html = f'{digest_html}<div class="news-filter">{filter_buttons}</div>{news_blocks}'

    # 테이블 행
    table_rows = ""
    for r in reversed(data):
        result_cls = {"승": "win", "패": "lose", "무": "draw", "취소": "cancel"}.get(r["result"], "")
        result_txt = r["result"] or "-"
        table_rows += f"""
        <tr>
          <td>{r["date"]}</td>
          <td>{r["home_away"] or "-"}</td>
          <td>{r["opponent"] or "-"}</td>
          <td><span class="badge {result_cls}">{result_txt}</span></td>
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
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Malgun Gothic', sans-serif; background: #f0f2f5; color: #1a1a2e; }}

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
  .table-section {{ padding: 0 32px 32px; }}
  .table-layout {{ display: flex; gap: 16px; align-items: flex-start; }}
  .table-card {{
    flex: 1; min-width: 0;
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

  /* 뉴스 패널 */
  .news-panel {{
    width: 300px; flex-shrink: 0;
    background: white; border-radius: 12px; padding: 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,.07);
    max-height: 600px; overflow-y: auto;
  }}
  .news-panel h3 {{ font-size: 13px; color: #555; margin-bottom: 10px; font-weight: 600; }}
  .news-filter {{ display: flex; flex-wrap: wrap; gap: 5px; margin-bottom: 14px; }}
  .news-filter button {{
    font-size: 11px; padding: 3px 9px; border-radius: 10px; border: 1px solid #ddd;
    background: #f8f9fa; color: #555; cursor: pointer; transition: all .15s;
  }}
  .news-filter button.active {{ background: #002D72; color: white; border-color: #002D72; }}
  .news-date {{ font-size: 11px; font-weight: 700; color: #002D72; margin: 12px 0 6px; padding-bottom: 4px; border-bottom: 1px solid #eef; }}
  .news-date:first-child {{ margin-top: 0; }}
  .news-item {{ margin-bottom: 10px; }}
  .news-src {{ display: inline-block; font-size: 10px; font-weight: 700; padding: 1px 6px; border-radius: 8px; margin-bottom: 3px; }}
  .badge-news {{ background: #e3f2fd; color: #1565c0; }}
  .badge-cafe {{ background: #e8f5e9; color: #2e7d32; }}
  .news-title {{ font-size: 12px; color: #333; line-height: 1.4; }}
  .news-title a, a.news-title {{ color: #333; text-decoration: none; }}
  .news-title:hover, a.news-title:hover {{ color: #002D72; text-decoration: underline; }}
  .news-desc {{ font-size: 11px; color: #888; margin-top: 2px; line-height: 1.4; }}
  .news-empty {{ font-size: 12px; color: #aaa; text-align: center; padding: 20px 0; }}
  .digest-box {{
    background: #f0f4ff; border-left: 3px solid #002D72;
    border-radius: 6px; padding: 10px 12px; margin-bottom: 14px;
  }}
  .digest-box .digest-label {{ font-size: 10px; font-weight: 700; color: #002D72; margin-bottom: 5px; }}
  .digest-box .digest-text  {{ font-size: 12px; color: #333; line-height: 1.6; white-space: pre-line; }}
  .news-views {{ font-size: 10px; color: #aaa; margin-left: 4px; }}

  .badge {{ display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 700; }}
  .badge.win    {{ background: #e8f5e9; color: #2e7d32; }}
  .badge.lose   {{ background: #ffebee; color: #c62828; }}
  .badge.draw   {{ background: #e3f2fd; color: #1565c0; }}
  .badge.cancel {{ background: #f5f5f5; color: #757575; }}

  @media (max-width: 900px) {{
    .charts {{ grid-template-columns: 1fr; }}
    .kpi-row, .charts, .table-section {{ padding-left: 16px; padding-right: 16px; }}
    .table-layout {{ flex-direction: column; }}
    .news-panel {{ width: 100%; max-height: 400px; }}
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
    <div class="value">{fmt(total_off)}<span class="unit">원</span></div>
    <div class="sub">{date_range_label}</div>
  </div>
  <div class="kpi">
    <div class="label">ON 거래액 (누계)</div>
    <div class="value">{fmt(total_on)}<span class="unit">원</span></div>
    <div class="sub">{date_range_label}</div>
  </div>
  <div class="kpi">
    <div class="label">총 거래액 (누계)</div>
    <div class="value red">{fmt(total_total)}<span class="unit">원</span></div>
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
</div>

<div class="charts">
  <div class="chart-card">
    <h3>날짜별 거래액 추이</h3>
    <canvas id="trendChart" height="120"></canvas>
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
  <div class="table-layout">
    <div class="table-card">
      <h3>전체 데이터</h3>
      <table>
        <thead>
          <tr>
            <th>날짜</th><th>홈/어웨이</th><th>상대팀</th><th>결과</th>
            <th style="text-align:right">OFF거래액</th>
            <th style="text-align:right">ON거래액</th>
            <th style="text-align:right">합계</th>
          </tr>
        </thead>
        <tbody>{table_rows}</tbody>
      </table>
    </div>
    <div class="news-panel">
      <h3>날짜별 주요 이슈</h3>
      {news_html}
    </div>
  </div>
</div>

<script>
const OFF = '#002D72', ON = '#C8102E', TOT = '#6c757d';
const alpha = (c, a) => c + Math.round(a*255).toString(16).padStart(2,'0');

// ── 추이 차트 ──
new Chart(document.getElementById('trendChart'), {{
  type: 'line',
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

// 뉴스 날짜 필터
function filterNews(date, btn) {{
  document.querySelectorAll('.news-filter button').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  document.querySelectorAll('.news-group').forEach(g => {{
    g.style.display = (date === 'all' || g.dataset.date === date) ? '' : 'none';
  }});
}}
</script>
</body>
</html>"""


def main():
    print("데이터 조회 중...")
    data, news, digest = fetch_data()
    print(f"병합된 데이터: {len(data)}일 / 뉴스이슈: {len(news)}건")

    os.makedirs("dashboard", exist_ok=True)
    html = build_html(data, news, digest)
    with open("dashboard/index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("dashboard/index.html 생성 완료")


if __name__ == "__main__":
    main()
