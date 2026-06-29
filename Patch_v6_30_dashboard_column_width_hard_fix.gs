/**
 * LOTTEON v6.30 dashboard column width hard fix
 *
 * 문제:
 * - 대시보드 시트가 생성/갱신된 뒤 열너비 자동조정이 실제 보기 상태에 반영되지 않았습니다.
 * - v6.23 열너비 자동조정은 전체 시트 순차 처리용이라, 대시보드 재생성 직후 다시 좁은 열너비가 남을 수 있었습니다.
 *
 * 수정:
 * - 대시보드 전용 열너비 강제 정리 함수를 추가합니다.
 * - 대시보드 생성/빠른 갱신/표시서식 정리/열너비 자동조정 후 항상 대시보드 전용 폭을 재적용합니다.
 * - 헤더 기준이 아니라 데이터 성격 기준으로 폭을 고정합니다.
 * - 요약 영역은 B/C/K열 중심으로 넓게, 브랜드 표 영역은 숫자 열을 좁게, 필터/메모 열은 넓게 정리합니다.
 */

var LOTTEON_PATCH_V630_DASHBOARD_COLUMN_WIDTH_HARD_FIX_LOADED = true;

var __baseRebuildDashboardSingleSource_v630 = typeof rebuildDashboardSingleSource_v628_ === 'function' ? rebuildDashboardSingleSource_v628_ : null;
var __baseHardFixDashboardDisplay_v630 = typeof hardFixDashboardDisplay_v625_ === 'function' ? hardFixDashboardDisplay_v625_ : null;
var __baseApplyDisplayStandardsOnlyFast_v630 = typeof applyDisplayStandardsOnlyFast_v623 === 'function' ? applyDisplayStandardsOnlyFast_v623 : null;
var __baseApplyFastOutputSheetFormatting_v630 = typeof applyFastOutputSheetFormatting_v620_ === 'function' ? applyFastOutputSheetFormatting_v620_ : null;
var __baseRunColumnWidthStep_v630 = typeof runColumnWidthAutoAdjustStep_v623 === 'function' ? runColumnWidthAutoAdjustStep_v623 : null;

rebuildDashboardSingleSource_v628_ = function(salesAgg, filterAgg) {
  var result = __baseRebuildDashboardSingleSource_v630 ? __baseRebuildDashboardSingleSource_v630.apply(this, arguments) : null;
  hardFixDashboardDisplay_v630_();
  return result;
};

hardFixDashboardDisplay_v625_ = function() {
  var result = __baseHardFixDashboardDisplay_v630 ? __baseHardFixDashboardDisplay_v630.apply(this, arguments) : null;
  hardFixDashboardDisplay_v630_();
  return result || { ok: true };
};

applyDisplayStandardsOnlyFast_v623 = function() {
  var result = __baseApplyDisplayStandardsOnlyFast_v630 ? __baseApplyDisplayStandardsOnlyFast_v630.apply(this, arguments) : null;
  hardFixDashboardDisplay_v630_();
  try { SpreadsheetApp.getUi().alert('표시서식 정리 완료\n\n대시보드 전용 열너비까지 다시 적용했습니다.'); } catch (e) {}
  return result;
};

applyFastOutputSheetFormatting_v620_ = function() {
  var result = __baseApplyFastOutputSheetFormatting_v630 ? __baseApplyFastOutputSheetFormatting_v630.apply(this, arguments) : null;
  hardFixDashboardDisplay_v630_();
  return result || { ok: true };
};

runColumnWidthAutoAdjustStep_v623 = function() {
  var result = __baseRunColumnWidthStep_v630 ? __baseRunColumnWidthStep_v630.apply(this, arguments) : null;
  hardFixDashboardDisplay_v630_();
  return result;
};

function hardFixDashboardDisplay_v630_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!sheet || sheet.getLastRow() < 1) return { skipped: true };

  var lastRow = Math.min(sheet.getLastRow(), 260);
  var lastCol = Math.min(sheet.getLastColumn(), 18);

  try {
    applyDashboardColumnWidths_v630_(sheet);
    applyDashboardRowsAndAlignment_v630_(sheet, lastRow, lastCol);
    applyDashboardNumberFormats_v630_(sheet, lastRow, lastCol);
    try { sheet.setFrozenRows(1); } catch (e) {}
    try { log_('patch_v630_dashboard_width_fix', 'rows=' + lastRow + ' cols=' + lastCol); } catch (e) {}
  } catch (e) {
    try { log_('patch_v630_dashboard_width_fix_error', String(e && e.message ? e.message : e)); } catch (ignore) {}
  }
  return { ok: true };
}

function applyDashboardColumnWidths_v630_(sheet) {
  // 대시보드 현재 구조는 A:P 중심입니다. Q 이후는 사용하지 않는 경우가 많아 좁게 둡니다.
  var widths = {
    1: 58,   // 구분
    2: 128,  // 항목 / 브랜드명
    3: 150,  // 값1 / 대표검색필터명
    4: 116,  // 값2 / 계정ID
    5: 70,   // 값3 / 수집수
    6: 70,   // 값4 / 매출상품
    7: 70,   // 값5 / 주문
    8: 112,  // 값6 / 순수매출액
    9: 112,  // 값7 / 실제정산금액
    10: 132, // 값8 / 예상정산금액
    11: 112, // 값9 / 정산기준금액
    12: 112, // 값10 / 매입금액
    13: 108, // 값11 / 예상이익
    14: 86,  // 값12 / 이익률/날짜
    15: 82,  // 값13 / 미정산/미판매수
    16: 230, // 메모
    17: 60,
    18: 60
  };
  Object.keys(widths).forEach(function(k) {
    try { sheet.setColumnWidth(Number(k), widths[k]); } catch (e) {}
  });
}

function applyDashboardRowsAndAlignment_v630_(sheet, lastRow, lastCol) {
  if (lastRow < 1 || lastCol < 1) return;
  var range = sheet.getRange(1, 1, lastRow, lastCol);
  range
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  try { sheet.setRowHeights(1, lastRow, 22); } catch (e) {}

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headerRows = [];
  for (var r = 0; r < values.length; r++) {
    var first = String(values[r][0] || '').replace(/\n/g, '').trim();
    if (r === 0 || first === '구분') headerRows.push(r + 1);
  }

  // 기본 데이터 행
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol)
      .setFontWeight('normal')
      .setBackground(null)
      .setHorizontalAlignment('left')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  // 헤더/섹션 행
  headerRows.forEach(function(rowNo) {
    try {
      sheet.getRange(rowNo, 1, 1, lastCol)
        .setBackground('#d9eaf7')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      sheet.setRowHeight(rowNo, 34);
    } catch (e) {}
  });

  // 숫자 영역 우측 정렬. 요약 C열, 표 E~O열.
  try { sheet.getRange(2, 3, Math.max(1, Math.min(28, lastRow - 1)), 1).setHorizontalAlignment('right'); } catch (e) {}
  try { if (lastRow > 31) sheet.getRange(32, 5, lastRow - 31, Math.min(11, lastCol - 4)).setHorizontalAlignment('right'); } catch (e) {}

  // 텍스트 주요 열 좌측 정렬
  [1,2,3,4,16].forEach(function(col) {
    if (col <= lastCol && lastRow > 1) {
      try { sheet.getRange(2, col, lastRow - 1, 1).setHorizontalAlignment('left'); } catch (e) {}
    }
  });
}

function applyDashboardNumberFormats_v630_(sheet, lastRow, lastCol) {
  if (lastRow < 2) return;
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  // 요약 영역: B열 항목명 기준 C열 포맷
  var summaryMax = Math.min(30, lastRow);
  for (var r = 2; r <= summaryMax; r++) {
    var item = normalizeDashboardHeader_v630_(values[r - 1][1]);
    if (!item) continue;
    try {
      var cell = sheet.getRange(r, 3);
      if (/율/.test(item)) cell.setNumberFormat('0.00%');
      else if (/일자|날짜|시각|최근|최초/.test(item)) cell.setNumberFormat('@');
      else if (/행수|상품수|건수|브랜드수|주문수|수집수/.test(item)) cell.setNumberFormat('#,##0');
      else if (/매출|정산|매입|이익|수수료|부가세|공급/.test(item)) cell.setNumberFormat('#,##0');
    } catch (e) {}
  }

  // 표 영역: 각 헤더 행별로 포맷 적용
  for (var hr = 1; hr <= lastRow; hr++) {
    var first = String(values[hr - 1][0] || '').replace(/\n/g, '').trim();
    if (!(hr === 1 || first === '구분')) continue;
    var nextHeader = findNextDashboardHeaderRow_v630_(values, hr + 1);
    var endRow = nextHeader ? nextHeader - 1 : lastRow;
    if (endRow <= hr) continue;
    var rowCount = endRow - hr;
    for (var c = 1; c <= lastCol; c++) {
      var h = normalizeDashboardHeader_v630_(values[hr - 1][c - 1]);
      if (!h) continue;
      try {
        var rg = sheet.getRange(hr + 1, c, rowCount, 1);
        if (/율/.test(h)) rg.setNumberFormat('0.00%').setHorizontalAlignment('right');
        else if (/날짜|일자|최근수집일/.test(h)) rg.setNumberFormat('@').setHorizontalAlignment('center');
        else if (/금액|매출|정산|매입|이익|수수료|부가세|공급/.test(h) && !/상품수|건수|수량/.test(h)) rg.setNumberFormat('#,##0').setHorizontalAlignment('right');
        else if (/수집수|상품|주문|건수|수량|미판매수|30일초과/.test(h) && !/주문번호/.test(h)) rg.setNumberFormat('#,##0').setHorizontalAlignment('right');
        else rg.setNumberFormat('@').setHorizontalAlignment('left');
      } catch (e) {}
    }
  }
}

function findNextDashboardHeaderRow_v630_(values, fromRowNo) {
  for (var r = fromRowNo; r <= values.length; r++) {
    var first = String(values[r - 1][0] || '').replace(/\n/g, '').trim();
    if (first === '구분') return r;
  }
  return 0;
}

function normalizeDashboardHeader_v630_(v) {
  return String(v == null ? '' : v).replace(/\n/g, '').replace(/\s+/g, '').trim();
}
