/**
 * 중학교 2학년 반편성 프로그램 (Google Apps Script)
 *
 * 구성
 *  - Code.gs   : 스프레드시트 읽기/쓰기, 메뉴, 웹앱 진입점
 *  - Index.html: 반편성 화면(자동 편성 알고리즘 + 수동 조정)
 *
 * 시트
 *  - 학생명단   : 1학년 학생 정보 (입력용)
 *  - 반편성결과 : 편성 결과 (새반/새번호)
 *  - 반별요약   : 반별 인원·성비·평균 등
 */

var SHEET_STUDENTS = '학생명단';
var SHEET_RESULT = '반편성결과';
var SHEET_SUMMARY = '반별요약';

var HEADERS = ['학번', '이름', '성별', '성적', '이전반', '특수', '생활지도', '분리학생', '동반학생', '고정반', '비고'];

/* ---------- 진입점 ---------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏫 반편성')
    .addItem('반편성 프로그램 열기', 'openDialog')
    .addSeparator()
    .addItem('학생명단 양식 만들기', 'createTemplate')
    .addItem('예시 데이터 채우기 (150명)', 'fillSampleData')
    .addToUi();
}

/** 웹앱으로 배포했을 때 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('중학교 2학년 반편성')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 스프레드시트 안에서 대화상자로 열기 */
function openDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Index').setWidth(1500).setHeight(900);
  SpreadsheetApp.getUi().showModelessDialog(html, '중학교 2학년 반편성');
}

/* ---------- 양식 / 예시 데이터 ---------- */

function getOrCreateSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function createTemplate() {
  var sh = getOrCreateSheet_(SHEET_STUDENTS);
  if (sh.getLastRow() > 0) {
    var ui = SpreadsheetApp.getUi();
    var r = ui.alert('학생명단 시트에 내용이 있습니다. 머리글만 다시 쓸까요?', ui.ButtonSet.YES_NO);
    if (r !== ui.Button.YES) return;
  }
  writeHeader_(sh);
  sh.getRange(2, 1).setNote(
    '학번: 1학년 학번 (예: 10312)\n성별: 남/여\n성적: 숫자(점수)\n이전반: 1학년 반 (비우면 학번에서 자동 추출)\n' +
    '특수/생활지도: Y 표시\n분리학생/동반학생: 학번 또는 이름, 쉼표로 구분\n고정반: 반드시 배정할 2학년 반 번호'
  );
}

function writeHeader_(sh) {
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, HEADERS.length, 90);
  sh.setColumnWidth(8, 140);
  sh.setColumnWidth(9, 140);
  sh.setColumnWidth(11, 200);
}

function fillSampleData() {
  var sh = getOrCreateSheet_(SHEET_STUDENTS);
  sh.clear();
  writeHeader_(sh);
  var last = '김이박최정강조윤장임한오서신권황안송류홍'.split('');
  var first = ['민준', '서준', '도윤', '예준', '시우', '하준', '주원', '지호', '지후', '준우', '서연', '서윤', '지우', '서현', '하은',
    '하윤', '민서', '지유', '윤서', '채원', '현우', '건우', '우진', '선우', '연우', '수아', '지민', '다은', '은서', '예은'];
  var male = first.slice(0, 10).concat(first.slice(20, 25));
  var female = first.slice(10, 20).concat(first.slice(25));
  var rows = [];
  for (var c = 1; c <= 5; c++) {
    for (var n = 1; n <= 30; n++) {
      var g = n <= 15 ? '남' : '여';
      var pool = g === '남' ? male : female;
      var fn = pool[Math.floor(Math.random() * pool.length)];
      var name = last[Math.floor(Math.random() * last.length)] + fn;
      var score = Math.round(Math.max(20, Math.min(100, 70 + 15 * gauss_())) * 10) / 10;
      rows.push([String(10000 + c * 100 + n), name, g, score, c,
        Math.random() < 0.03 ? 'Y' : '', Math.random() < 0.08 ? 'Y' : '', '', '', '', '']);
    }
  }
  // 분리/동반 예시
  rows[2][7] = rows[40][0];
  rows[10][7] = rows[11][0] + ',' + rows[70][0];
  rows[95][8] = rows[96][0];
  sh.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  sh.getRange(2, 1, rows.length, 1).setNumberFormat('@');
}

function gauss_() {
  var u = 1 - Math.random(), v = Math.random();
  return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
}

/* ---------- 클라이언트 호출 API ---------- */

/** 학생명단 시트를 읽어 객체 배열로 반환 */
function getStudents() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STUDENTS);
  if (!sh || sh.getLastRow() < 2) {
    throw new Error('"' + SHEET_STUDENTS + '" 시트에 데이터가 없습니다. 메뉴 > 반편성 > 학생명단 양식 만들기 후 입력하세요.');
  }
  var values = sh.getDataRange().getDisplayValues();
  var head = values[0].map(function (h) { return String(h).trim(); });
  var idx = {};
  HEADERS.forEach(function (h) { idx[h] = head.indexOf(h); });
  if (idx['이름'] < 0) throw new Error('"이름" 열을 찾을 수 없습니다.');

  var get = function (row, key) { return idx[key] >= 0 ? String(row[idx[key]] || '').trim() : ''; };
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var name = get(row, '이름');
    if (!name) continue;
    out.push({
      id: get(row, '학번') || ('R' + (r + 1)),
      name: name,
      gender: get(row, '성별'),
      score: get(row, '성적'),
      prev: get(row, '이전반'),
      special: get(row, '특수'),
      care: get(row, '생활지도'),
      separate: get(row, '분리학생'),
      together: get(row, '동반학생'),
      fixed: get(row, '고정반'),
      note: get(row, '비고')
    });
  }
  return out;
}

/**
 * 결과 저장
 * @param {{rows: Array<Array>, summary: Array<Array>, header: string[], summaryHeader: string[]}} payload
 */
function saveResult(payload) {
  var sh = getOrCreateSheet_(SHEET_RESULT);
  sh.clear();
  var header = payload.header;
  var rows = payload.rows;
  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff').setHorizontalAlignment('center');
  if (rows.length) {
    sh.getRange(2, 1, rows.length, header.length).setNumberFormat('@').setValues(rows.map(function (r) {
      return r.map(function (v) { return v === null || v === undefined ? '' : String(v); });
    }));
  }
  sh.setFrozenRows(1);
  // 반이 바뀔 때마다 줄무늬
  var colors = [];
  var prev = null, toggle = false;
  rows.forEach(function (r) {
    if (r[0] !== prev) { toggle = !toggle; prev = r[0]; }
    colors.push(new Array(header.length).fill(toggle ? '#ffffff' : '#eef4fb'));
  });
  if (rows.length) sh.getRange(2, 1, rows.length, header.length).setBackgrounds(colors);

  var sm = getOrCreateSheet_(SHEET_SUMMARY);
  sm.clear();
  var sh2 = payload.summaryHeader;
  sm.getRange(1, 1, 1, sh2.length).setValues([sh2])
    .setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff').setHorizontalAlignment('center');
  if (payload.summary.length) sm.getRange(2, 1, payload.summary.length, sh2.length).setValues(payload.summary);
  sm.setFrozenRows(1);

  return { ok: true, count: rows.length, savedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') };
}

/** 저장된 결과(학번 → 새반)를 불러오기 */
function loadSavedResult() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULT);
  if (!sh || sh.getLastRow() < 2) return null;
  var values = sh.getDataRange().getDisplayValues();
  var head = values[0];
  var iCls = head.indexOf('새반'), iId = head.indexOf('학번');
  if (iCls < 0 || iId < 0) return null;
  var map = {};
  for (var r = 1; r < values.length; r++) {
    var id = String(values[r][iId]).trim();
    var cls = parseInt(values[r][iCls], 10);
    if (id && cls) map[id] = cls;
  }
  return map;
}
