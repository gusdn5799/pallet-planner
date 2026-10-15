/**
 * 혼적 파레트 적재 플래너 → 구글시트 '파랫트 적재' 시트 연동
 *
 * 플래너에서 "전체 지시서 이미지 복사"를 누르면 A4 이미지(파레트 4장씩)가 이 웹앱으로 전송되고,
 * '파랫트 적재' 시트의 기존 이미지를 모두 지운 뒤 새 이미지를 A4 크기로 위에서부터 붙입니다.
 *
 * 설치: README.md의 "구글시트 연동" 참고
 */

const SHEET_NAME = '파랫트 적재';
const LIST_SHEET = '쿠팡 적재리스트';   // 파레트별 쿠팡 적재리스트 양식을 쓰는 시트 (없으면 새로 만듦)
// A4 가로를 96dpi 화면 기준 픽셀로 (297×210mm, 인쇄 여백만큼 높이를 조금 줄임)
const PAGE_W = 1123;
const PAGE_H = 780;

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

      // 2) 칸 크기를 A4에 맞춤: A열 = A4 폭, 한 쪽 = A4 높이
      //    시트 이미지 제한(1장 최대 100만 픽셀) 때문에 플래너가 한 쪽을 tilesPerPage장(위→아래)으로 나눠 보냄
      const per = Math.max(1, Number(body.tilesPerPage) || 1);
      const n = body.pages.length;
      const pages = Math.ceil(n / per);
      sh.setColumnWidth(1, PAGE_W);
      if (sh.getMaxRows() < n) sh.insertRowsAfter(sh.getMaxRows(), n - sh.getMaxRows());
      const tileH = [];
      for (let i = 0; i < n; i++) {
        const k = i % per;                                   // 쪽 안에서 몇 번째 조각인지
        const h = Math.round(PAGE_H * (k + 1) / per) - Math.round(PAGE_H * k / per);
        tileH.push(h);
        sh.setRowHeight(i + 1, h);
      }
      // 이전에 더 많은 쪽을 올렸던 경우 남은 행 높이를 원래대로 (빈 쪽 인쇄 방지)
      if (sh.getMaxRows() > n) sh.setRowHeights(n + 1, sh.getMaxRows() - n, 21);

      // 3) 새 이미지 등록 (A1, A2, … 에 조각 하나씩)
      body.pages.forEach(function (dataUrl, i) {
        const m = String(dataUrl).match(/^data:(image\/(?:png|jpeg));base64,(.+)$/);
        if (!m) throw new Error((i + 1) + '번째 이미지 형식이 올바르지 않습니다');
        const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'pallet-' + (i + 1));
        const img = sh.insertImage(blob, 1, i + 1);
        img.setWidth(PAGE_W).setHeight(tileH[i]);
      });
      // 메모는 PDF 인쇄 때 별도 페이지로 붙으므로 남기지 않음 (예전 버전이 남긴 메모도 지움)
      sh.getRange(1, 1, sh.getMaxRows(), 1).clearNote();
      SpreadsheetApp.flush();

      // 4) 쿠팡 적재리스트 양식 (파레트별 칸)
      let listUrl = '';
      if (Array.isArray(body.list) && body.list.length) {
        const ls = writeCoupangList_(ss, body.list);
        listUrl = ss.getUrl() + '#gid=' + ls.getSheetId();
      }

      const gid = sh.getSheetId();
      const pdfUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf'
        + '&gid=' + gid + '&size=A4&portrait=false&fitw=true&gridlines=false&printtitle=false'
        + '&sheetnames=false&pagenum=UNDEFINED&horizontal_alignment=CENTER'
        + '&top_margin=0.15&bottom_margin=0.15&left_margin=0.15&right_margin=0.15';
      return json_({ ok: true, pages: pages, sheetUrl: ss.getUrl() + '#gid=' + gid, pdfUrl: pdfUrl, listUrl: listUrl });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/**
 * 쿠팡 파렛트 적재리스트 양식을 파레트마다 하나씩 아래로 이어서 작성.
 * 채우는 칸: 박스수량( 00 BOX ), SKU NAME, BOX 수량. 나머지는 양식 그대로(빈칸).
 * list = [{ total: 20, items: [{ name: '…', box: 3 }, …] }, …]
 */
function writeCoupangList_(ss, list) {
  const ls = ss.getSheetByName(LIST_SHEET) || ss.insertSheet(LIST_SHEET);
  ls.getRange(1, 1, ls.getMaxRows(), ls.getMaxColumns()).breakApart();
  ls.clear();
  const widths = [70, 110, 470, 100, 100, 110];
  widths.forEach(function (w, i) { ls.setColumnWidth(i + 1, w); });

  let r = 1;
  list.forEach(function (p) {
    const n = Math.max(10, p.items.length);              // 양식은 10줄, 품목이 더 많으면 늘림
    const need = r + 6 + n + 1;
    if (ls.getMaxRows() < need) ls.insertRowsAfter(ls.getMaxRows(), need - ls.getMaxRows());

    const head = [
      ['쿠팡 파렛트 적재리스트(혼적 시 필수작성)'],
      ['파렛트(0개) – (NO.) / 박스수량 ( ' + p.total + ' BOX )'],
      ['입고예정일자 ( 00월 00일 ) / 입고처 ( 인천17 센터 )'],
      ['업체명 ( ㈜올계농업회사법인 ) /'],
      ['발주번호 (0)']
    ];
    head.forEach(function (row, i) {
      ls.getRange(r + i, 1, 1, 6).merge().setValue(row[0]);
    });
    ls.getRange(r, 1, 1, 6).setFontSize(16).setFontWeight('bold').setFontLine('underline').setHorizontalAlignment('center');
    ls.getRange(r + 1, 1, 4, 6).setFontSize(13).setFontWeight('bold').setHorizontalAlignment('left');

    const th = r + 5;
    ls.getRange(th, 1, 1, 6).setValues([['NO', 'SKU No', 'SKU NAME', 'BOX 수량', '수량', '소비기한']])
      .setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
    const rows = [];
    for (let i = 0; i < n; i++) {
      const it = p.items[i];
      rows.push([i + 1, '', it ? it.name : '', it ? it.box : '', '', '']);
    }
    ls.getRange(th + 1, 1, n, 6).setValues(rows).setFontSize(10).setHorizontalAlignment('center');
    ls.getRange(th + 1, 1, n, 1).setFontWeight('bold');
    ls.getRange(r, 1, 6 + n, 6).setBorder(true, true, true, true, true, true).setVerticalAlignment('middle');
    ls.getRange(r + 1, 1, 4, 6).setBorder(null, null, null, null, false, false);   // 머리글 4줄은 줄 구분선 없이
    ls.getRange(r + 1, 1, 4, 6).setBorder(true, true, true, true, null, null);
    r += 6 + n + 2;                                        // 파레트 사이 한 줄 띄움
  });
  return ls;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
