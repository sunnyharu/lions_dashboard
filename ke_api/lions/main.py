"""라이온즈 대시보드 API.

  GET  /api/dashboard  사내 시트를 읽어 만든 대시보드 HTML (캐시 후 반환)
  POST /api/note       특이사항 저장 {date, note}
  POST /api/upload     온라인 상품별 업로드 {key, rows}
  GET  /health

대시보드 HTML 은 요청 시점에 시트에서 생성하므로 실데이터가 레포에 남지 않는다.
"""
import os
import threading
import time

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse, PlainTextResponse

from . import generate_dashboard as gd
from .sheets_write import upload_online_rows, write_note

CACHE_TTL = int(os.environ.get("DASHBOARD_CACHE_TTL", "600"))
UPLOAD_KEY = os.environ.get("UPLOAD_KEY", "")

# 사내 배포의 보안 헤더(CSP)가 외부 스크립트와 외부 fetch 를 막으므로 같은 출처로 바꿔 끼운다.
REWRITES = [
    ("https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js", "/vendor/chart.umd.min.js"),
    ("https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js", "/vendor/xlsx.full.min.js"),
    ("https://cdn.jsdelivr.net/gh/projectnoonnu/noonfonts_2304-2@1.0/", "/vendor/fonts/"),
    ("https://lions-note.freelywind222.workers.dev", "/api/note"),
    # 시트 반영이 즉시라 업로드 후 3분 대기할 필요가 없다
    ("약 2~3분 후 대시보드가 자동으로 갱신됩니다.", "잠시 후 새로고침됩니다."),
    ("setTimeout(() => location.reload(), 3 * 60 * 1000);", "setTimeout(() => location.reload(), 4000);"),
    ("업데이트 중... (약 1~2분 소요)", "업데이트 중..."),
]

app = FastAPI()
_cache = {"html": None, "at": 0.0}
_lock = threading.Lock()


def _build() -> str:
    data, news, digest, off, on = gd.fetch_data()
    html = gd.build_html(data, news, digest, off, on, apps_script_url="/api/upload")
    for old, new in REWRITES:
        html = html.replace(old, new)
    return html


def _refresh():
    with _lock:
        _cache["html"] = _build()
        _cache["at"] = time.time()


def _invalidate():
    _cache["at"] = 0.0


@app.on_event("startup")
def _warm():
    threading.Thread(target=_refresh, daemon=True).start()


@app.get("/health")
def health():
    return {"ok": True, "cached": _cache["html"] is not None}


@app.get("/api/dashboard", response_class=HTMLResponse)
def dashboard():
    if _cache["html"] is None:
        _refresh()                                   # 첫 요청은 생성될 때까지 기다린다
    elif time.time() - _cache["at"] > CACHE_TTL and not _lock.locked():
        threading.Thread(target=_refresh, daemon=True).start()   # 오래된 캐시는 일단 주고 뒤에서 갱신
    return HTMLResponse(_cache["html"])


@app.post("/api/note")
async def note(req: Request):
    body = await req.json()
    date, text = str(body.get("date", "")).strip(), str(body.get("note", "")).strip()
    if not date or not text:
        return PlainTextResponse("날짜와 내용을 모두 입력해주세요.", status_code=400)
    write_note(date, text)
    _invalidate()
    return {"ok": True}


@app.post("/api/upload")
async def upload(req: Request):
    # Apps Script 와 같은 응답 형식 {ok, inserted | error} — 화면 코드가 그대로 읽는다
    body = await req.json()
    if not UPLOAD_KEY or body.get("key") != UPLOAD_KEY:
        return JSONResponse({"ok": False, "error": "비밀번호가 틀렸습니다."})
    rows = body.get("rows") or []
    if not rows:
        return JSONResponse({"ok": False, "error": "데이터가 없습니다."})
    try:
        result = upload_online_rows(rows)
    except Exception as e:
        return JSONResponse({"ok": False, "error": str(e)[:200]})
    _invalidate()
    return {"ok": True, **result}
