/**
 * LOTTEON v6.23 column width batch + display format patch
 *
 * 목적:
 * 1) 열너비 자동조정을 별도 메뉴/함수로 분리합니다.
 * 2) 시간초과 방지를 위해 1회 실행당 일부 시트만 처리하는 이어실행 구조로 동작합니다.
 * 3) 열너비는 헤더가 아니라 아래 데이터 기준으로 계산합니다.
 * 4) 헤더가 가려질 수 있는 출력 시트는 의미 단위 줄바꿈 헤더로 표시합니다.
 * 5) 금액 컬럼의 통화기호(₩)를 제거하고 #,##0 기준으로 통일합니다.
 * 6) 날짜 표시는 MM/dd 기준으로 통일합니다.
 *
 * 운영:
 * - LOTTEON 자동화 메뉴에 추가되면 "열너비 자동조정 이어실행"을 완료될 때까지 반복 실행합니다.
 * - 메뉴가 보이지 않으면 Apps Script 함수 선택에서 runColumnWidthAutoAdjustBatch_v623 실행합니다.
 */

var LOTTEON_PATCH_V623_COLUMN_WIDTH_BATCH_DISPLAY_FORMAT_LOADED = true;
var LOTTEON_COLUMN_WIDTH_BATCH_STATE_KEY_V623 = 'LOTTEON_COLUMN_WIDTH_BATCH_STATE_V623';
var LOTTEON_COLUMN_WIDTH_BATCH_SIZE_V623 = 2;
var LOTTEON_COLUMN_WIDTH_TIME_LIMIT_MS_V623 = 90000;

var __baseOnOpen_v623 = typeof onOpen === 'function' ? onOpen : null;
onOpen = function() {
  if (__baseOnOpen_v623) {
    try { __baseOnOpen_v623.apply(this, arguments); } catch (e) {}
  }
  addColumnWidthMenu_v623_();
};

function addColumnWidthMenu_v623_() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('LOTTEON 서식')
      .addItem('열너비 자동조정 이어실행', 'runColumnWidthAutoAdjustBatch_v623')
      .addItem('열너비 자동조정 초기화', 'resetColumnWidthAutoAdjustBatch_v623')
      .addToUi();
  } catch (e) {}
}

function runColumnWidthAutoAdjustBatch_v623() {
  var ui = SpreadsheetApp.getUi();
  var started = Date.now();
  var ss = SpreadsheetApp.getActive();
  var props = PropertiesService.getScriptProperties();
  var state = loadColumnWidthState_v623_(props, ss);
  var sheets = getColumnWidthTargetSheets_v623_(ss);
  if (!sheets.length) {
    ui.alert('열너비 자동조정 대상 시트가 없습니다.');
    return;
  }

  var processed = [];
  var startIndex = Math.min(state.nextIndex || 0, sheets.length);
  var i = startIndex;
  var count = 0;

  for (; i < sheets.length; i++) {
    if (count >= LOTTEON_COLUMN_WIDTH_BATCH_SIZE_V623) break;
    if (Date.now() - started > LOTTEON_COLUMN_WIDTH_TIME_LIMIT_MS_V623) break;
    var sheet = sheets[i];
    formatSheetColumnsByData_v623_(sheet);
    processed.push(sheet.getName());
    count++;
  }

  var done = i >= sheets.length;
  if (done) {
    props.deleteProperty(LOTTEON_COLUMN_WIDTH_BATCH_STATE_KEY_V623);
    ui.alert(
      '열너비 자동조정 완료\n\n' +
      '이번 처리 시트:\n- ' + (processed.length ? processed.join('\n- ') : '(없음)') + '\n\n' +
      '전체 대상 시트 ' + sheets.length + '개 처리가 완료됐습니다.'
    );
  } else {
    props.setProperty(LOTTEON_COLUMN_WIDTH_BATCH_STATE_KEY_V623, JSON.stringify({ nextIndex: i, total: sheets.length, updatedAt: new Date().toISOString() }));
    ui.alert(
      '열너비 자동조정 일부 완료\n\n' +
      '이번 처리 시트:\n- ' + (processed.length ? processed.join('\n- ') : '(없음)') + '\n\n' +
      '진행: ' + i + ' / ' + sheets.length + '\n' +
      '완료될 때까지 같은 메뉴를 다시 실행하세요.'
    );
  }
}

function resetColumnWidthAutoAdjustBatch_v623() {
  PropertiesService.getScriptProperties().deleteProperty(LOTTEON_COLUMN_WIDTH_BATCH_STATE_KEY_V623);
  SpreadsheetApp.getUi().alert('열너비 자동조정 이어실행 상태를 초기화했습니다.');
}

function loadColumnWidthState_v623_(props, ss) {
  try {
    var raw = props.getProperty(LOTTEON_COLUMN_WIDTH_BATCH_STATE_KEY_V623);
    if (!raw) return { nextIndex: 0 };
    var state = JSON.parse(raw);
    if (!state || typeof state.nextIndex !== 'number') return { nextIndex: 0 };
    return state;
  } catch (e) {
    return { nextIndex: 0 };
  }
}

function getColumnWidthTargetSheets_v623_(ss) {
  var priorityNames = [
    CONFIG.SHEETS.DASHBOARD,
    '브랜드별_마진율',
    '미정산_쿠팡계정별',
    '부가세_신고자료',
    '부가세_상품별',
    '부가세_고객별',
    '부가세_주문번호별',
    CONFIG.SHEETS.RETRANSMIT_LOG,
    CONFIG.SHEETS.FILTERS,
    CONFIG.SHEETS.SALES_CLEAN,
    CONFIG.SHEETS.SALES_IN
  ];
  var result = [];
  var seen = {};
  priorityNames.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (sh && !seen[sh.getSheetId()]) {
      result.push(sh);
      seen[sh.getSheetId()] = true;
    }
  });
  ss.getSheets().forEach(function(sh) {
    if (!seen[sh.getSheetId()]) {
      result.push(sh);
      seen[sh.getSheetId()] = true;
    }
  });
  return result;
}

function formatSheetColumnsByData_v623_(sheet) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var outputSheet = isOutputSheet_v623_(sheet.getName());
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  var rawHeaders = headerRange.getValues()[0].map(function(h) { return String(h || '').trim(); });

  if (outputSheet) {
    var displayHeaders = rawHeaders.map(function(h) { return headerDisplayText_v623_(h); });
    headerRange.setValues([displayHeaders]);
  }

  try {
    sheet.getRange(1, 1, Math.min(lastRow, 500), lastCol)
      .setFontFamily('Arial')
      .setFontSize(10)
      .setVerticalAlignment('middle');
    headerRange
      .setBackground('#d9eaf7')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    sheet.setFrozenRows(1);
  } catch (e) {}

  var sampleRowCount = Math.min(Math.max(lastRow - 1, 0), 300);
  var dataValues = sampleRowCount > 0 ? sheet.getRange(2, 1, sampleRowCount, lastCol).getDisplayValues() : [];

  for (var c = 1; c <= lastCol; c++) {
    var header = rawHeaders[c - 1];
    var profile = inferColumnProfile_v623_(header, dataValues, c - 1);
    var width = estimateWidthFromData_v623_(header, dataValues, c - 1, profile);
    try {
      sheet.setColumnWidth(c, width);
      applyColumnFormat_v623_(sheet, c, lastRow, profile);
      if (profile.longText) {
        sheet.getRange(1, c, Math.min(lastRow, 600), 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        sheet.getRange(1, c, 1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      }
    } catch (e) {}
  }

  // 대시보드 중간 헤더 행도 다시 스타일 적용
  if (sheet.getName() === CONFIG.SHEETS.DASHBOARD) {
    styleDashboardSectionHeaders_v623_(sheet);
  }

  try {
    sheet.setRowHeight(1, 42);
  } catch (e) {}
}

function inferColumnProfile_v623_(header, dataValues, idx) {
  var h = normalizeHeaderForType_v623_(header);
  var sample = [];
  for (var r = 0; r < dataValues.length; r++) {
    var v = String(dataValues[r][idx] || '').trim();
    if (v) sample.push(v);
    if (sample.length >= 30) break;
  }
  var joined = sample.join(' ');
  var p = { money: false, percent: false, count: false, date: false, idText: false, longText: false };
  if (/금액|매출|정산|매입|이익|수수료|부가세|공급|배송비|원가|가격|단가/.test(h)) p.money = true;
  if (/율|비율|rate|%/.test(h)) p.percent = true;
  if (/수량|건수|주문수|상품수|고객수|수집수|전송수|미정산건수|초과건수|판매수량|매출상품|주문$/.test(h)) p.count = true;
  if (/날짜|일자|일시|주문일|수집일|생성일|등록일|갱신시각|최근|최초|년|월|일$/.test(h)) p.date = true;
  if (/주문번호|상품번호|계정ID|필터ID|상품ID|코드|번호/.test(h)) p.idText = true;
  if (/상품명|메모|비고|사유|주소|요청|필터명|대표검색필터명/.test(h)) p.longText = true;

  // 값 기준 보정: 헤더가 애매해도 데이터가 금액/퍼센트/날짜처럼 보이면 반영
  if (!p.money && sample.length && sample.filter(function(v) { return /^₩?[-]?[\d,]+$/.test(v); }).length >= Math.max(3, sample.length * 0.7) && /금|매출|정산|매입|이익|수수료|부가세|공급|원가/.test(joined + h)) p.money = true;
  if (!p.percent && sample.some(function(v) { return /%$/.test(v); })) p.percent = true;
  if (!p.date && sample.some(function(v) { return /^(20\d{2}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2})/.test(v); })) p.date = true;

  if (p.money) { p.percent = false; p.count = false; }
  if (p.percent) { p.money = false; p.count = false; }
  return p;
}

function estimateWidthFromData_v623_(header, dataValues, idx, profile) {
  var lengths = [];
  for (var r = 0; r < dataValues.length; r++) {
    var text = String(dataValues[r][idx] || '').trim();
    if (!text) continue;
    // 헤더가 아니라 아래 데이터 기준: 긴 텍스트는 상한만 적용
    lengths.push(Math.min(text.length, profile.longText ? 45 : 28));
  }
  var maxLen = lengths.length ? Math.max.apply(null, lengths) : 8;
  var width = 55 + maxLen * 7;
  if (profile.money) width = Math.max(width, 105);
  if (profile.percent) width = Math.max(width, 80);
  if (profile.count) width = Math.max(width, 75);
  if (profile.idText) width = Math.max(width, 115);
  if (profile.date) width = Math.max(width, 70);
  if (profile.longText) width = Math.max(width, 160);

  var max = profile.longText ? 360 : (profile.money ? 135 : (profile.idText ? 180 : 150));
  var min = profile.date ? 65 : 55;
  return Math.max(min, Math.min(max, width));
}

function applyColumnFormat_v623_(sheet, col, lastRow, profile) {
  if (lastRow < 2) return;
  var range = sheet.getRange(2, col, lastRow - 1, 1);
  if (profile.money) range.setNumberFormat('#,##0');
  else if (profile.percent) range.setNumberFormat('0.00%');
  else if (profile.date) range.setNumberFormat('MM/dd');
  else if (profile.idText) range.setNumberFormat('@');
  else if (profile.count) range.setNumberFormat('#,##0');
}

function styleDashboardSectionHeaders_v623_(sheet) {
  var lastRow = Math.min(sheet.getLastRow(), 160);
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1) return;
  var values = sheet.getRange(1, 1, lastRow, Math.min(lastCol, 16)).getValues();
  for (var r = 0; r < values.length; r++) {
    if (r === 0 || String(values[r][0] || '').trim() === '구분') {
      sheet.getRange(r + 1, 1, 1, lastCol)
        .setBackground('#d9eaf7')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    }
  }
}

function isOutputSheet_v623_(name) {
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

function headerDisplayText_v623_(header) {
  var h = String(header || '').trim();
  var map = {
    '대표검색필터명': '대표\n검색필터명',
    '예상정산금액(미정산)': '예상정산금액\n(미정산)',
    '마켓수수료/비용': '마켓수수료\n/비용',
    '부가세반영예상이익': '부가세반영\n예상이익',
    '부가세반영이익률': '부가세반영\n이익률',
    '30일초과미정산건수': '30일초과\n미정산건수',
    '30일초과건수': '30일초과\n건수',
    '30일초과 주문번호': '30일초과\n주문번호',
    '미정산주문건수': '미정산\n주문건수',
    '미정산건수': '미정산\n건수',
    '정산기준금액': '정산기준\n금액',
    '실제정산금액': '실제정산\n금액',
    '순수매출액': '순수\n매출액',
    '총매출액': '총\n매출액',
    '매입금액': '매입\n금액',
    '예상이익률': '예상\n이익률',
    '납부예상부가세': '납부예상\n부가세',
    '매출공급가액': '매출\n공급가액',
    '매출부가세': '매출\n부가세',
    '매입공급가액': '매입\n공급가액',
    '매입부가세': '매입\n부가세'
  };
  if (map[h]) return map[h];
  return h;
}

function normalizeHeaderForType_v623_(header) {
  return String(header || '').replace(/\n/g, '').replace(/\s+/g, '').trim();
}

// v6.20 빠른 서식 함수도 새 기준으로 교체: ⑤ 이후에는 전체가 아니라 운영 출력 시트만 빠르게 적용합니다.
applyFastOutputSheetFormatting_v620_ = function() {
  var ss = SpreadsheetApp.getActive();
  [CONFIG.SHEETS.DASHBOARD, '브랜드별_마진율', '미정산_쿠팡계정별'].forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (sh) formatSheetColumnsByData_v623_(sh);
  });
  return { formatted: 3, mode: 'v6.23 output fast' };
};
