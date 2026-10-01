# ke_api — 사내 ke-ops 배포용 API 프로토타입

사내 계정 시트를 읽어 대시보드 HTML 을 요청 시점에 만들어 준다(실데이터를 레포에 커밋하지 않음).
Cloudflare Worker(특이사항)·Apps Script(업로드)를 같은 출처 API 로 대체한다.

- `lions/main.py` — FastAPI: `/api/dashboard`(10분 캐시), `/api/note`, `/api/upload`, `/health`
- `lions/sheets_write.py` — 특이사항·온라인 업로드 시트 쓰기
- `lions/generate_dashboard.py` — 레포 루트 `generate_dashboard.py` 를 복사하고
  `SPREADSHEET_ID` 를 `os.environ.get("SPREADSHEET_ID", "<사내 시트 ID>")` 로 바꿔 쓴다

사내 보안 헤더(CSP)가 외부 스크립트·글꼴·fetch 를 막으므로 `main.REWRITES` 로 같은 출처 경로로 치환한다.
`/vendor/chart.umd.min.js`, `/vendor/xlsx.full.min.js`, `/vendor/fonts/KBO-Dia-Gothic_*.woff` 는
web 서비스의 `public/vendor/` 에 두어야 한다.

로컬 확인: `UPLOAD_KEY=... uvicorn lions.main:app --port 8091` (2026-10-01 사내 시트로 렌더 검증, 11.5MB·첫 생성 13초)
