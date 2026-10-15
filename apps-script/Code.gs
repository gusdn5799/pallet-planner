/**
 * 혼적 파레트 적재 플래너 → 구글시트 '파랫트 적재' 시트 연동
 *
 * 플래너에서 "전체 지시서 이미지 복사"를 누르면 A4 이미지(파레트 4장씩)가 이 웹앱으로 전송되고,
 * '파랫트 적재' 시트의 기존 이미지를 모두 지운 뒤 새 이미지를 A4 크기로 위에서부터 붙입니다.
 *
 * 설치: README.md의 "구글시트 연동" 참고
 */

const SHEET_NAME = '파랫트 적재';
// A4 세로를 96dpi 화면 기준 픽셀로 (210×297mm)
const PAGE_W = 794;
const PAGE_H = 1123;

/** (선택) 비밀키를 쓰려면 한 번 실행. 사내용으로 키 없이 쓸 때는 실행하지 않아도 됩니다. */
function setup() {
  const key = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('TOKEN', key);
  Logger.log('비밀키: ' + key + '  ← 플래너의 "구글시트 연동 설정"에 붙여넣으세요.');
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
    if (token && body.token !== token) return json_({ ok: false, error: '비밀키가 맞지 않습니다' });
    if (!Array.isArray(body.pages) || !body.pages.length) return json_({ ok: false, error: '이미지가 없습니다' });

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

      // 1) 이전 이미지 전부 삭제
      sh.getImages().forEach(function (img) { img.remove(); });

      // 2) 칸 크기를 A4에 맞춤: A열 = A4 폭, 쪽마다 한 행 = A4 높이
      sh.setColumnWidth(1, PAGE_W);
      const n = body.pages.length;
      if (sh.getMaxRows() < n) sh.insertRowsAfter(sh.getMaxRows(), n - sh.getMaxRows());
      for (let i = 0; i < n; i++) sh.setRowHeight(i + 1, PAGE_H);

      // 3) 새 이미지 등록 (A1, A2, … 에 한 쪽씩)
      body.pages.forEach(function (dataUrl, i) {
        const m = String(dataUrl).match(/^data:(image\/(?:png|jpeg));base64,(.+)$/);
        if (!m) throw new Error((i + 1) + '쪽 이미지 형식이 올바르지 않습니다');
        const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'pallet-' + (i + 1));
        const img = sh.insertImage(blob, 1, i + 1);
        img.setWidth(PAGE_W).setHeight(PAGE_H);
      });
      sh.getRange('A1').setNote('플래너에서 등록: ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'));
      SpreadsheetApp.flush();

      const gid = sh.getSheetId();
      const pdfUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf'
        + '&gid=' + gid + '&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false'
        + '&sheetnames=false&pagenum=UNDEFINED&horizontal_alignment=CENTER'
        + '&top_margin=0.2&bottom_margin=0.2&left_margin=0.2&right_margin=0.2';
      return json_({ ok: true, pages: n, sheetUrl: ss.getUrl() + '#gid=' + gid, pdfUrl: pdfUrl });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
