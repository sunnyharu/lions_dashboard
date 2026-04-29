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
        try: off = int(str(off).replace(",", ""))
        except: off = 0
        try: on  = int(str(on).replace(",", ""))
        except: on  = 0
        sales[date] = {"off": off, "on": on}

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

    return merged


def build_html(data: list) -> str:
    game_days = [r for r in data if r["result"]]

    def avg(lst): return int(sum(lst) / len(lst)) if lst else 0
    def fmt(n):   return f"{n:,}" if n else "-"

    total_off   = sum(r["off"]   for r in data if r["off"])
    total_on    = sum(r["on"]    for r in data if r["on"])
    total_total = total_off + total_on

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

    home_avg_off  = avg([r["off"]   for r in home_rows])
    away_avg_off  = avg([r["off"]   for r in away_rows])
    home_avg_on   = avg([r["on"]    for r in home_rows])
    away_avg_on   = avg([r["on"]    for r in away_rows])

    win_avg_off   = avg([r["off"]   for r in win_rows])
    lose_avg_off  = avg([r["off"]   for r in lose_rows])
    win_avg_on    = avg([r["on"]    for r in win_rows])
    lose_avg_on   = avg([r["on"]    for r in lose_rows])

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
  .kpi .sub   {{ font-size: 11px; color: #aaa; margin-top: 4px; }}

  /* 차트 그리드 */
  .charts {{ display: grid; grid-template-columns: 2fr 1fr 1fr; gap: 16px; padding: 20px 32px; }}
  .chart-card {{
    background: white; border-radius: 12px; padding: 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,.07);
  }}
  .chart-card h3 {{ font-size: 13px; color: #555; margin-bottom: 16px; font-weight: 600; }}

  /* 테이블 */
  .table-section {{ padding: 0 32px 32px; }}
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

  .badge {{ display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 700; }}
  .badge.win    {{ background: #e8f5e9; color: #2e7d32; }}
  .badge.lose   {{ background: #ffebee; color: #c62828; }}
  .badge.draw   {{ background: #e3f2fd; color: #1565c0; }}
  .badge.cancel {{ background: #f5f5f5; color: #757575; }}

  @media (max-width: 900px) {{
    .charts {{ grid-template-columns: 1fr; }}
    .kpi-row, .charts, .table-section {{ padding-left: 16px; padding-right: 16px; }}
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
    <div class="value">{fmt(total_off)}</div>
    <div class="sub">원</div>
  </div>
  <div class="kpi">
    <div class="label">ON 거래액 (누계)</div>
    <div class="value">{fmt(total_on)}</div>
    <div class="sub">원</div>
  </div>
  <div class="kpi">
    <div class="label">총 거래액 (누계)</div>
    <div class="value red">{fmt(total_total)}</div>
    <div class="sub">원</div>
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
    <canvas id="trendChart" height="220"></canvas>
  </div>
  <div class="chart-card">
    <h3>홈 / 어웨이 평균 거래액</h3>
    <canvas id="haChart" height="220"></canvas>
  </div>
  <div class="chart-card">
    <h3>경기 결과별 평균 거래액</h3>
    <canvas id="resultChart" height="220"></canvas>
  </div>
</div>

<div class="table-section">
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

// ── 홈/어웨이 ──
new Chart(document.getElementById('haChart'), {{
  type: 'bar',
  data: {{
    labels: ['홈', '어웨이'],
    datasets: [
      {{ label: 'OFF거래액', data: [{home_avg_off}, {away_avg_off}], backgroundColor: alpha(OFF, .8) }},
      {{ label: 'ON거래액',  data: [{home_avg_on},  {away_avg_on}],  backgroundColor: alpha(ON,  .8) }},
    ]
  }},
  options: {{
    responsive: true, plugins: {{ legend: {{ position: 'top' }} }},
    scales: {{ y: {{ ticks: {{ callback: v => v.toLocaleString() }} }} }}
  }}
}});

// ── 결과별 ──
new Chart(document.getElementById('resultChart'), {{
  type: 'bar',
  data: {{
    labels: ['승', '패'],
    datasets: [
      {{ label: 'OFF거래액', data: [{win_avg_off}, {lose_avg_off}], backgroundColor: alpha(OFF, .8) }},
      {{ label: 'ON거래액',  data: [{win_avg_on},  {lose_avg_on}],  backgroundColor: alpha(ON,  .8) }},
    ]
  }},
  options: {{
    responsive: true, plugins: {{ legend: {{ position: 'top' }} }},
    scales: {{ y: {{ ticks: {{ callback: v => v.toLocaleString() }} }} }}
  }}
}});
</script>
</body>
</html>"""


def main():
    print("데이터 조회 중...")
    data = fetch_data()
    print(f"병합된 데이터: {len(data)}일")

    os.makedirs("dashboard", exist_ok=True)
    html = build_html(data)
    with open("dashboard/index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("dashboard/index.html 생성 완료")


if __name__ == "__main__":
    main()
