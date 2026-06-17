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

    // 기존 데이터 읽기
    const allVals = ws.getDataRange().getValues();
    const hasHeader = allVals.length > 0 && String(allVals[0][0]) === SHEET_HEADER[0];
    const existingRows = hasHeader ? allVals.slice(1) : [];

    // 업로드 날짜에 해당하지 않는 기존 행만 유지
    const kept = existingRows.filter(r => {
      const d = String(r[0] || '').trim().replace(/\./g, '-');
      return !uploadDates.has(d);
    });

    // 새 행
    const newRows = rows.map(r => SHEET_HEADER.map(col => String(r[col] ?? '')));

    // 시트 전체 초기화 후 한번에 재작성
    ws.clearContents();
    const writeData = [SHEET_HEADER, ...kept, ...newRows];
    ws.getRange(1, 1, writeData.length, SHEET_HEADER.length)
      .setValues(writeData)
      .setNumberFormat('@');

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
