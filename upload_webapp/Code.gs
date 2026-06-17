// ── 설정 ──────────────────────────────────────────────
// Script Properties에 저장:
//   SECRET_KEY   : 업로드 비밀번호
//   GITHUB_TOKEN : repo workflow 트리거용 PAT (repo + workflow scope)
const SPREADSHEET_ID = '1ylkJlnm1ykfazJXV65HKt5cH5IXudWEeKBKLt_SzplU';
const SHEET_NAME     = '상품별매출(on)';
const GITHUB_REPO    = 'sunnyharu/lions_dashboard';
const GITHUB_WORKFLOW = 'update_dashboard.yml';

const SHEET_HEADER = ['판매일자','상품ID','상품명','바코드','skucode','사이즈','선수명','판매단가','판매수량','실판매금액'];

function doPost(e) {
  const res = ContentService.createTextOutput();
  res.setMimeType(ContentService.MimeType.JSON);

  try {
    const props = PropertiesService.getScriptProperties();
    const SECRET_KEY   = props.getProperty('SECRET_KEY');
    const GITHUB_TOKEN = props.getProperty('GITHUB_TOKEN');

    const body = JSON.parse(e.postData.contents);

    // 비밀번호 검증
    if (!SECRET_KEY || body.key !== SECRET_KEY) {
      res.setContent(JSON.stringify({ ok: false, error: '비밀번호가 틀렸습니다.' }));
      return res;
    }

    const rows = body.rows;
    if (!rows || rows.length === 0) {
      res.setContent(JSON.stringify({ ok: false, error: '데이터가 없습니다.' }));
      return res;
    }

    // 업로드 데이터의 날짜 목록
    const uploadDates = new Set(rows.map(r => String(r['판매일자'] || '').trim()));

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ws = ss.getSheetByName(SHEET_NAME);

    // 헤더 확인
    const allVals = ws.getDataRange().getValues();
    if (allVals.length === 0 || String(allVals[0][0]) !== SHEET_HEADER[0]) {
      ws.clearContents();
      ws.appendRow(SHEET_HEADER);
    }

    // 기존 데이터에서 업로드 날짜 해당 행 삭제 (뒤에서부터)
    const data = ws.getDataRange().getValues();
    const dateColIdx = 0; // 판매일자는 첫 번째 컬럼
    for (let i = data.length - 1; i >= 1; i--) {
      const d = String(data[i][dateColIdx] || '').trim().replace(/\./g, '-');
      if (uploadDates.has(d) || uploadDates.has(d.replace(/-/g, '.'))) {
        ws.deleteRow(i + 1); // 1-based
      }
    }

    // 새 행 추가
    const newRows = rows.map(r => SHEET_HEADER.map(col => r[col] ?? ''));
    if (newRows.length > 0) {
      ws.getRange(ws.getLastRow() + 1, 1, newRows.length, SHEET_HEADER.length)
        .setValues(newRows)
        .setNumberFormat('@'); // 바코드 등 텍스트 유지
    }

    // GitHub Actions 트리거
    let triggered = false;
    if (GITHUB_TOKEN) {
      const apiUrl = `https://api.github.com/repos/${GITHUB_REPO}/actions/workflows/${GITHUB_WORKFLOW}/dispatches`;
      const resp = UrlFetchApp.fetch(apiUrl, {
        method: 'post',
        headers: {
          'Authorization': `Bearer ${GITHUB_TOKEN}`,
          'Accept': 'application/vnd.github+json',
          'Content-Type': 'application/json',
        },
        payload: JSON.stringify({ ref: 'main' }),
        muteHttpExceptions: true,
      });
      triggered = resp.getResponseCode() === 204;
    }

    res.setContent(JSON.stringify({
      ok: true,
      inserted: rows.length,
      dates: [...uploadDates],
      dashboardTriggered: triggered,
    }));
  } catch (err) {
    res.setContent(JSON.stringify({ ok: false, error: err.message }));
  }

  return res;
}
