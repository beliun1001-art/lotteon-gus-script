/**
 * LOTTEON v6.23 column width batch / display standard patch
 *
 * 목적:
 * 1) 열너비 자동조정 메뉴를 별도 메뉴로 분리합니다.
 * 2) 시간초과 방지를 위해 1회 실행당 일부 시트만 처리하는 이어실행 구조로 만듭니다.
 * 3) 열너비는 헤더가 아니라 실제 데이터 표시값 기준으로 계산합니다.
 * 4) 헤더가 가려질 수 있는 경우 헤더는 줄바꿈 표시합니다.
 * 5) 금액 서식에서 통화기호(₩)를 제거합니다.
 * 6) 출력용 시트의 날짜 표시는 MM/dd, 즉 00/00 기준으로 통일합니다.
 *
 * 주의:
 * - 매출데이터_붙여넣기, 매출데이터_정리 같은 원본/중간 시트의 헤더 값은 변경하지 않습니다.
 * - 출력용 시트만 긴 헤더 텍스트를 의미 단위로 줄바꿈하고, 년/월/일 3열 구조를 날짜 1열(MM/dd)로 정리합니다.
 */

var LOTTEON_PATCH_V623_COLUMN_WIDTH_BATCH_DISPLAY_STANDARD_LOADED = true;
var LOTTEON_COLWIDTH_V623_NEXT_INDEX = 'LOTTEON_COLWIDTH_V623_NEXT_INDEX';
var LOTTEON_COLWIDTH_V623_LAST_RESULT = 'LOTTEON_COLWIDTH_V623_LAST_RESULT';
var LOTTEON_COLWIDTH_V623_BATCH_SIZE = 2;
var LOTTEON_COLWIDTH_V623_SAMPLE_ROWS = 900;

var __baseOnOpen_v623 = typeof onOpen === 'function' ? onOpen : null;
onOpen = function() {
  if (__baseOnOpen_v623) {
    try { __baseOnOpen_v623.apply(this, arguments); } catch (e) {}
  }
  addLotteonFormatMenu_v623_();
};

// 금액 포맷 공통 함수 override: 통화기호 제거
formatMoneyColumns_v611_ = function(sheet, cols) {
  var lastRow = sheet && sheet.getLastRow ? sheet.getLastRow() : 0;
  if (!sheet || lastRow < 2) return;
  (cols || []).forEach(function(col) {
    try { sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('#,##0'); } catch (e) {}
  });
};

// ⑤ 빠른 갱신에서 호출되는 v6.20 서식 함수도 통화기호 없는 빠른 버전으로 대체
applyFastOutputSheetFormatting_v620_ = function() {
  var ss = SpreadsheetApp.getActive();
  getFastFormatSheetNames_v623_().forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) applyDisplayStandardsFast_v623_(sheet);
  });
  try { log_('patch_v623_fast_display_standard', 'done'); } catch (e) {}
  return { ok: true };
};

fastFormatOneSheet_v620_ = function(sheet) {
  return applyDisplayStandardsFast_v623_(sheet);
};

function addLotteonFormatMenu_v623_() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('LOTTEON 서식')
    .addItem('열너비 자동조정 시작/이어실행', 'runColumnWidthAutoAdjustStep_v623')
    .addItem('열너비 자동조정 상태 확인', 'showColumnWidthAutoAdjustStatus_v623')
    .addItem('열너비 자동조정 초기화', 'resetColumnWidthAutoAdjust_v623')
    .addSeparator()
    .addItem('표시서식만 빠른 정리', 'applyDisplayStandardsOnlyFast_v623')
    .addToUi();
}

function resetColumnWidthAutoAdjust_v623() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(LOTTEON_COLWIDTH_V623_NEXT_INDEX);
  props.deleteProperty(LOTTEON_COLWIDTH_V623_LAST_RESULT);
  SpreadsheetApp.getUi().alert('열너비 자동조정 진행 상태를 초기화했습니다.');
}

function showColumnWidthAutoAdjustStatus_v623() {
  var ss = SpreadsheetApp.getActive();
  var props = PropertiesService.getScriptProperties();
  var sheets = getColumnWidthTargetSheets_v623_(ss);
  var idx = toNumber_(props.getProperty(LOTTEON_COLWIDTH_V623_NEXT_INDEX));
  var last = props.getProperty(LOTTEON_COLWIDTH_V623_LAST_RESULT) || '';
  SpreadsheetApp.getUi().alert(
    '열너비 자동조정 상태\n\n' +
    '대상 시트 수: ' + sheets.length + '\n' +
    '다음 처리 위치: ' + Math.min(idx + 1, sheets.length + 1) + ' / ' + sheets.length + '\n' +
    '최근 결과: ' + (last || '없음')
  );
}

function applyDisplayStandardsOnlyFast_v623() {
  var ss = SpreadsheetApp.getActive();
  var count = 0;
  getFastFormatSheetNames_v623_().forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    applyDisplayStandardsFast_v623_(sheet);
    count++;
  });
  SpreadsheetApp.getUi().alert('표시서식 빠른 정리 완료\n\n처리 시트: ' + count + '개\n금액: 통화기호 없음\n날짜: MM/dd 기준');
}

function runColumnWidthAutoAdjustStep_v623() {
  var ui = SpreadsheetApp.getUi();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    ui.alert('다른 작업이 실행 중입니다. 잠시 후 다시 실행해주세요.');
    return;
  }

  var started = Date.now();
  try {
    var ss = SpreadsheetApp.getActive();
    var props = PropertiesService.getScriptProperties();
    var sheets = getColumnWidthTargetSheets_v623_(ss);
    var startIndex = toNumber_(props.getProperty(LOTTEON_COLWIDTH_V623_NEXT_INDEX));
    if (startIndex < 0 || startIndex >= sheets.length) startIndex = 0;

    var endIndex = Math.min(startIndex + LOTTEON_COLWIDTH_V623_BATCH_SIZE, sheets.length);
    var processed = [];
    for (var i = startIndex; i < endIndex; i++) {
      var sheet = sheets[i];
      normalizeOutputSheetBeforeWidth_v623_(sheet);
      adjustSheetColumnWidthsByData_v623_(sheet);
      processed.push(sheet.getName());
    }

    var elapsedSec = Math.round((Date.now() - started) / 1000);
    var done = endIndex >= sheets.length;
    if (done) {
      props.deleteProperty(LOTTEON_COLWIDTH_V623_NEXT_INDEX);
    } else {
      props.setProperty(LOTTEON_COLWIDTH_V623_NEXT_INDEX, String(endIndex));
    }

    var result = '처리=' + processed.join(', ') + ' / 위치=' + endIndex + '/' + sheets.length + ' / 소요초=' + elapsedSec;
    props.setProperty(LOTTEON_COLWIDTH_V623_LAST_RESULT, result);
    try { log_('column_width_batch_v623', result); } catch (e) {}

    ui.alert(
      (done ? '열너비 자동조정 완료' : '열너비 자동조정 일부 완료') + '\n\n' +
      '처리 시트:\n- ' + processed.join('\n- ') + '\n\n' +
      '진행: ' + endIndex + ' / ' + sheets.length + '\n' +
      '소요초: ' + elapsedSec + '\n\n' +
      (done ? '전체 대상 시트 처리가 끝났습니다.' : '남은 시트가 있습니다. 같은 메뉴를 다시 실행하세요.')
    );
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    try { log_('column_width_batch_v623_error', msg); } catch (ignore) {}
    ui.alert('열너비 자동조정 중 오류가 발생했습니다.\n\n' + msg);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

function getColumnWidthTargetSheets_v623_(ss) {
  var preferred = [
    CONFIG.SHEETS.DASHBOARD,
    '브랜드별_마진율',
    '미정산_쿠팡계정별',
    CONFIG.SHEETS.RETRANSMIT_LOG,
    '부가세_신고자료',
    '부가세_상품별',
    '부가세_고객별',
    '부가세_주문번호별',
    CONFIG.SHEETS.FILTERS,
    CONFIG.SHEETS.SALES_IN,
    CONFIG.SHEETS.SALES_CLEAN,
    CONFIG.SHEETS.BRAND_SUMMARY
  ];
  var added = {};
  var result = [];
  preferred.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet && !added[name]) { result.push(sheet); added[name] = true; }
  });
  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (!added[name]) { result.push(sheet); added[name] = true; }
  });
  return result;
}

function getFastFormatSheetNames_v623_() {
  return [
    CONFIG.SHEETS.DASHBOARD,
    '브랜드별_마진율',
    '미정산_쿠팡계정별',
    CONFIG.SHEETS.RETRANSMIT_LOG,
    '부가세_신고자료',
    '부가세_상품별',
    '부가세_고객별',
    '부가세_주문번호별'
  ];
}

function normalizeOutputSheetBeforeWidth_v623_(sheet) {
  if (!isGeneratedOutputSheet_v623_(sheet.getName())) return;
  combineYearMonthDayToDateColumn_v623_(sheet);
  normalizeOutputSheetMoneyAndDates_v623_(sheet);
  applySemanticHeaderBreaks_v623_(sheet);
}

function isGeneratedOutputSheet_v623_(name) {
  return [
    CONFIG.SHEETS.DASHBOARD,
    CONFIG.SHEETS.RETRANSMIT_LOG,
    '브랜드별_마진율',
    '부가세_신고자료',
    '부가세_상품별',
    '부가세_고객별',
    '부가세_주문번호별',
    '미정산_쿠팡계정별'
  ].indexOf(name) >= 0;
}

function combineYearMonthDayToDateColumn_v623_(sheet) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 3) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var h0 = String(values[0][0] || '').replace(/\n/g, '').trim();
  var h1 = String(values[0][1] || '').replace(/\n/g, '').trim();
  var h2 = String(values[0][2] || '').replace(/\n/g, '').trim();
  if (!(h0 === '년' && h1 === '월' && h2 === '일')) return;

  var out = [];
  for (var r = 0; r < values.length; r++) {
    if (r === 0) {
      out.push(['날짜'].concat(values[r].slice(3)));
    } else {
      out.push([mmddFromParts_v623_(values[r][0], values[r][1], values[r][2])].concat(values[r].slice(3)));
    }
  }
  sheet.clearContents();
  sheet.getRange(1, 1, out.length, out[0].length).setValues(out);
}

function normalizeOutputSheetMoneyAndDates_v623_(sheet) {
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || '').replace(/\n/g, '').trim(); });
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var changed = false;

  for (var c = 0; c < header.length; c++) {
    var h = header[c];
    var money = isMoneyHeader_v623_(h);
    var date = isDateHeader_v623_(h);
    if (!money && !date) continue;

    for (var r = 0; r < values.length; r++) {
      var before = values[r][c];
      var after = before;
      if (money) after = normalizeMoneyValue_v623_(before);
      if (date) after = normalizeDateDisplayValue_v623_(after);
      if (after !== before) { values[r][c] = after; changed = true; }
    }
  }

  if (changed) sheet.getRange(2, 1, lastRow - 1, lastCol).setValues(values);
}

function applySemanticHeaderBreaks_v623_(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return;
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var changed = false;
  var out = headers.map(function(h) {
    var raw = String(h || '').replace(/\n/g, '').trim();
    var next = semanticHeaderBreak_v623_(raw);
    if (next !== h) changed = true;
    return next;
  });
  if (changed) sheet.getRange(1, 1, 1, lastCol).setValues([out]);
}

function applyDisplayStandardsFast_v623_(sheet) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return { skipped: true };
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (isGeneratedOutputSheet_v623_(sheet.getName())) applySemanticHeaderBreaks_v623_(sheet);
  applyNumberDateFormatsNoCurrency_v623_(sheet, Math.min(lastRow, 5000));
  try {
    sheet.getRange(1, 1, 1, lastCol)
      .setFontWeight('bold')
      .setBackground('#d9eaf7')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    sheet.setFrozenRows(1);
  } catch (e) {}
  return { ok: true };
}

function adjustSheetColumnWidthsByData_v623_(sheet) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var sampleRows = Math.min(Math.max(lastRow - 1, 0), LOTTEON_COLWIDTH_V623_SAMPLE_ROWS);
  var dataValues = sampleRows > 0 ? sheet.getRange(2, 1, sampleRows, lastCol).getDisplayValues() : [];
  var headerValues = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];

  applyNumberDateFormatsNoCurrency_v623_(sheet, lastRow);

  for (var c = 0; c < lastCol; c++) {
    var header = String(headerValues[c] || '').replace(/\n/g, '').trim();
    var maxLen = 0;
    for (var r = 0; r < dataValues.length; r++) {
      var text = String(dataValues[r][c] || '').replace(/\s+/g, ' ').trim();
      if (text.length > maxLen) maxLen = text.length;
    }
    if (!maxLen) maxLen = Math.min(header.length, 12);
    var width = widthFromDataLength_v623_(header, maxLen);
    try { sheet.setColumnWidth(c + 1, width); } catch (e) {}
  }

  try {
    sheet.getRange(1, 1, 1, lastCol)
      .setFontWeight('bold')
      .setBackground('#d9eaf7')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    if (lastRow > 1) {
      sheet.getRange(2, 1, Math.min(lastRow - 1, 1000), lastCol).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    }
    sheet.setFrozenRows(1);
  } catch (e) {}

  if (sheet.getName() === CONFIG.SHEETS.DASHBOARD && typeof styleDashboardSectionHeadersFast_v620_ === 'function') {
    styleDashboardSectionHeadersFast_v620_(sheet, Math.min(lastRow, 160), Math.min(lastCol, 24));
  }
}

function applyNumberDateFormatsNoCurrency_v623_(sheet, rowLimit) {
  if (!sheet || sheet.getLastRow() < 2) return;
  var lastRow = Math.min(sheet.getLastRow(), rowLimit || sheet.getLastRow());
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || '').replace(/\n/g, '').trim(); });
  var rows = Math.max(lastRow - 1, 1);

  for (var c = 0; c < headers.length; c++) {
    var h = headers[c];
    try {
      var range = sheet.getRange(2, c + 1, rows, 1);
      if (isMoneyHeader_v623_(h)) range.setNumberFormat('#,##0');
      else if (/율|rate|%/i.test(h)) range.setNumberFormat('0.00%');
      else if (isDateHeader_v623_(h)) range.setNumberFormat('MM/dd');
      else if (/수집수|상품수|주문|건수|수량|고객수|미판매수|경과일|count|qty/i.test(h) && !/주문번호/.test(h)) range.setNumberFormat('#,##0');
      else if (/주문번호|상품번호|계정ID|ID$/i.test(h)) range.setNumberFormat('@');
    } catch (e) {}
  }
}

function isMoneyHeader_v623_(header) {
  var h = String(header || '').replace(/\n/g, '').trim();
  if (/율|rate|%/.test(h)) return false;
  if (/금액|매출|정산|매입|이익|수수료|부가세|공급가액|공급대가|배송비|원가|비용|amount|price|cost|fee|vat/i.test(h)) return true;
  return false;
}

function isDateHeader_v623_(header) {
  var h = String(header || '').replace(/\n/g, '').trim();
  if (h === '년' || h === '월' || h === '일') return false;
  return /날짜|일자|주문일|결제일|수집일|생성일|등록일|전송일|갱신일|완료일|최근|최초|date/i.test(h);
}

function normalizeMoneyValue_v623_(value) {
  if (typeof value === 'number') return value;
  var text = String(value == null ? '' : value).trim();
  if (!text) return value;
  if (text.indexOf('₩') < 0 && text.indexOf(',') < 0) return value;
  var normalized = text.replace(/₩/g, '').replace(/,/g, '').trim();
  if (/^-?\d+(\.\d+)?$/.test(normalized)) return Number(normalized);
  return text.replace(/₩/g, '').trim();
}

function normalizeDateDisplayValue_v623_(value) {
  if (!value) return value;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, CONFIG.TIMEZONE, 'MM/dd');
  }
  var text = String(value).trim();
  var m = text.match(/20\d{2}[-.\/](\d{1,2})[-.\/](\d{1,2})/);
  if (m) return ('0' + Number(m[1])).slice(-2) + '/' + ('0' + Number(m[2])).slice(-2);
  var md = text.match(/^(\d{1,2})[-.\/](\d{1,2})$/);
  if (md) return ('0' + Number(md[1])).slice(-2) + '/' + ('0' + Number(md[2])).slice(-2);
  return value;
}

function mmddFromParts_v623_(year, month, day) {
  if (!month || !day) return '';
  return ('0' + Number(month)).slice(-2) + '/' + ('0' + Number(day)).slice(-2);
}

function semanticHeaderBreak_v623_(header) {
  var h = String(header || '').trim();
  var map = {
    '대표검색필터명': '대표\n검색필터명',
    '예상정산금액(미정산)': '예상정산금액\n(미정산)',
    '마켓수수료/비용': '마켓수수료\n/비용',
    '부가세반영예상이익': '부가세반영\n예상이익',
    '부가세반영이익률': '부가세반영\n이익률',
    '정산기준금액': '정산기준\n금액',
    '실제정산금액': '실제정산\n금액',
    '예상정산금액': '예상정산\n금액',
    '순수매출액': '순수\n매출액',
    '총매출액': '총\n매출액',
    '취소/반품매출': '취소/반품\n매출',
    '매입금액': '매입\n금액',
    '예상이익': '예상\n이익',
    '예상이익률': '예상\n이익률',
    '미정산건수': '미정산\n건수',
    '30일초과건수': '30일초과\n건수',
    '30일초과 주문번호': '30일초과\n주문번호',
    '납부예상부가세': '납부예상\n부가세',
    '매출공급가액': '매출\n공급가액',
    '매입공급가액': '매입\n공급가액',
    '매출부가세': '매출\n부가세',
    '매입부가세': '매입\n부가세'
  };
  if (map[h]) return map[h];
  if (h.length >= 9) {
    h = h.replace(/(금액|매출액|정산금액|부가세|이익률|주문건수|상품수|필터명)$/g, '\n$1');
  }
  return h;
}

function widthFromDataLength_v623_(header, dataLen) {
  var h = String(header || '');
  var max = maxColumnWidthForHeader_v623_(h);
  var min = minColumnWidthForHeader_v623_(h);
  var width = Math.round(dataLen * 8.5 + 26);
  if (/상품명/.test(h)) width = Math.round(dataLen * 7.2 + 28);
  if (/메모|비고|사유/.test(h)) width = Math.round(dataLen * 7.0 + 28);
  return Math.max(min, Math.min(max, width));
}

function minColumnWidthForHeader_v623_(header) {
  var h = String(header || '');
  if (/날짜|주문일|월|일/.test(h)) return 64;
  if (/율/.test(h)) return 76;
  if (/금액|매출|정산|매입|이익|수수료|부가세/.test(h)) return 98;
  if (/계정ID/.test(h)) return 105;
  if (/브랜드명/.test(h)) return 105;
  return 58;
}

function maxColumnWidthForHeader_v623_(header) {
  var h = String(header || '');
  if (/상품명/.test(h)) return 360;
  if (/메모|비고|사유/.test(h)) return 260;
  if (/대표검색필터명|필터명/.test(h)) return 220;
  if (/주문번호/.test(h)) return 180;
  if (/고객명/.test(h)) return 130;
  if (/브랜드명/.test(h)) return 145;
  if (/상품번호/.test(h)) return 135;
  if (/계정ID/.test(h)) return 120;
  if (/금액|매출|정산|매입|이익|수수료|부가세|공급/.test(h)) return 135;
  return 155;
}
