/**
 * 서부기초학력지원청 서부기초학력지원센터 · 2026학년도 2학기 찾아가는 학습클리닉
 * 보강 신청 게시판 백엔드 (Google Apps Script)
 *
 * 이 파일은 "스프레드시트에 연결된" Apps Script 프로젝트에 붙여 넣어 사용합니다.
 * 배포 방법은 같은 폴더의 README.md 를 참고하세요.
 *
 *  - doPost : 신청 내용을 시트에 한 줄 추가
 *  - doGet  : 신청 내역을 JSON(또는 JSONP)으로 반환
 */

var SHEET_NAME = '보강신청';
var HEADERS = ['접수시각', '강사명', '학교', '학생', '원래수업일', '원래시간', '보강날짜', '보강시간', '사유', '상태'];

/** 시트를 가져오고, 없으면 머리글까지 만들어 준다. */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  // 날짜·시간 열이 자동으로 날짜 값으로 바뀌지 않도록 텍스트 서식으로 둔다. (A, E ~ H)
  sh.getRange('A:A').setNumberFormat('@');
  sh.getRange('E:H').setNumberFormat('@');
  return sh;
}

/** 시트가 날짜/시간을 Date 객체로 바꿔 놓았을 때를 대비한 정규화 */
function normText_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  return String(v);
}
/** 접수시각은 날짜와 시각을 함께 남긴다. */
function normStamp_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  }
  return String(v);
}
function normTime_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, 'Asia/Seoul', 'HH:mm');
  }
  return String(v);
}

function json_(obj, callback) {
  var body = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + body + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}

/** 신청 내역 조회 */
function doGet(e) {
  var callback = (e && e.parameter && e.parameter.callback) || '';
  try {
    var sh = getSheet_();
    var last = sh.getLastRow();
    var items = [];
    if (last > 1) {
      var rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        if (!r[1] && !r[4] && !r[6]) continue; // 빈 행 건너뛰기
        items.push({
          time:     normStamp_(r[0]),
          teacher:  normText_(r[1]),
          school:   normText_(r[2]),
          student:  normText_(r[3]),
          fromDate: normText_(r[4]),
          fromTime: normTime_(r[5]),
          toDate:   normText_(r[6]),
          toTime:   normTime_(r[7]),
          reason:   normText_(r[8]),
          status:   normText_(r[9]) || '신청'
        });
      }
    }
    items.reverse(); // 최근 신청순
    return json_({ ok: true, count: items.length, items: items }, callback);
  } catch (err) {
    return json_({ ok: false, error: String(err) }, callback);
  }
}

/** 보강 신청 접수 */
function doPost(e) {
  try {
    var p = (e && e.parameter) || {};
    var teacher  = String(p.teacher  || '').trim();
    var fromDate = String(p.fromDate || '').trim();
    var toDate   = String(p.toDate   || '').trim();
    var toTime   = String(p.toTime   || '').trim();

    if (!teacher || !fromDate || !toDate || !toTime) {
      return json_({ ok: false, error: '필수 항목이 누락되었습니다.' });
    }

    getSheet_().appendRow([
      Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
      teacher,
      String(p.school   || '').trim(),
      String(p.student  || '').trim(),
      fromDate,
      String(p.fromTime || '').trim(),
      toDate,
      toTime,
      String(p.reason   || '').trim(),
      '신청'
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
