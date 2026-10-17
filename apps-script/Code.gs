/**
 * 1923 종로 장날 · 제출 우체통
 *
 * 학생이 게임에서 「선생님께 제출」을 누르면 이 시트에 한 줄씩 쌓입니다.
 * 설치: 구글 시트에서 [확장 프로그램] → [Apps Script] → 이 코드를 통째로 붙여 넣기
 *       → [배포] → [새 배포] → 유형 "웹 앱", 실행 사용자 "나", 액세스 권한 "모든 사용자" → 배포
 *       → 나온 웹 앱 주소(…/exec)를 게임 index.html 의 SUBMIT_URL 에 넣기
 */

const SHEET_NAME = '제출';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    if (!data || !data['학번'] || !data['이름']) return json({ ok: false, error: 'missing id' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    // 머리글: 처음이면 만들고, 새 항목이 생기면 오른쪽에 덧붙입니다.
    let headers = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    if (!headers.length) headers = ['제출 시각'];
    Object.keys(data).forEach(k => { if (headers.indexOf(k) < 0) headers.push(k); });
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);

    const row = headers.map(h => h === '제출 시각' ? new Date() : clean(data[h]));
    sh.appendRow(row);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 주소가 살아 있는지 브라우저로 열어 확인할 때 씁니다.
function doGet() {
  return json({ ok: true, message: '우체통이 열려 있습니다.' });
}

// 학생 글이 수식으로 실행되지 않도록 막고, 너무 긴 글은 자릅니다.
function clean(v) {
  if (v === undefined || v === null) return '';
  let s = String(v).slice(0, 5000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
