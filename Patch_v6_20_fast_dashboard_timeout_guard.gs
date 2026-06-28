/**
 * LOTTEON v6.20 fast dashboard timeout guard patch
 *
 * 문제:
 * - v6.19에서 ⑤ 대시보드만 빠른 갱신 후 전체 시트를 대상으로 autoResizeColumns/긴 텍스트 축약을 수행하면서
 *   실행 시간이 길어져 최대 실행 시간 초과가 발생했습니다.
 *
 * 수정:
 * - ⑤ 대시보드만 빠른 갱신에서는 전체 시트 스캔/전체 셀 축약을 하지 않습니다.
 * - 대시보드/운영 출력 시트만 헤더 기반 폭으로 빠르게 정리합니다.
 * - 대량 행이 있는 부가세 상세 시트는 autoResize 대신 안전한 고정 폭을 적용합니다.
 * - 출력값 자체를 다시 훑으며 축약하지 않고, 화면 표시만 CLIP 처리합니다.
 *
 * 결과:
 * - ⑤ 메뉴는 시간초과 방지를 우선합니다.
 * - 전체 시트 세밀한 폭 조정은 이후 별도 배치/분할 메뉴로 분리하는 것이 안전합니다.
 */

var LOTTEON_PATCH_V620_FAST_DASHBOARD_TIMEOUT_GUARD_LOADED = true;

// v6.19/v6.18의 무거운 formatting 함수를 빠른 버전으로 대체합니다.
applyAllSheetsAutoWidthAndCompact_v619_ = function() {
  return applyFastOutputSheetFormatting_v620_();
};

applyWorkbookCompactFormatting_v618_ = function() {
  return applyFastOutputSheetFormatting_v620_();
};

autoWidthAndLimitSheet_v619_ = function(sheet) {
  return fastFormatOneSheet_v620_(sheet);
};

compactSheetDisplay_v618_ = function(sheet) {
  return fastFormatOneSheet_v620_(sheet);
};

// 시간초과 원인이 되는 전체 셀 값 재작성 축약은 ⑤에서는 하지 않습니다.
compactLongTextCells_v619_ = function(sheet) {
  return { skipped: true, reason: 'v6.20 time guard' };
};

compactLongTextCellsByHeader_v618_ = function(sheet) {
  return { skipped: true, reason: 'v6.20 time guard' };
};

function applyFastOutputSheetFormatting_v620_() {
  var ss = SpreadsheetApp.getActive();
  var names = [
    CONFIG.SHEETS.DASHBOARD,
    CONFIG.SHEETS.RETRANSMIT_LOG,
    '브랜드별_마진율',
    '부가세_신고자료',
    '부가세_상품별',
    '부가세_고객별',
    '부가세_주문번호별',
    '미정산_쿠팡계정별'
  ];

  var formatted = [];
  names.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    fastFormatOneSheet_v620_(sheet);
    formatted.push(name);
  });

  try { log_('patch_v620_fast_format', 'formatted=' + formatted.join(',')); } catch (e) {}
  return { formatted: formatted.length, sheets: formatted };
}

function fastFormatOneSheet_v620_(sheet) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return { skipped: true };

  var lastRow = sheet.getLastRow();
  var lastCol = Math.min(sheet.getLastColumn(), 24);
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || '').trim(); });
  var isLarge = lastRow > 250;

  try {
    sheet.getRange(1, 1, Math.min(lastRow, 250), lastCol)
      .setFontFamily('Arial')
      .setFontSize(10)
      .setVerticalAlignment('middle');
    sheet.getRange(1, 1, 1, lastCol)
      .setFontWeight('bold')
      .setBackground('#d9eaf7')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  } catch (e) {}

  // 작은 운영 시트는 autoResize를 허용하고, 큰 상세 시트는 고정 폭만 적용합니다.
  try {
    if (!isLarge) sheet.autoResizeColumns(1, lastCol);
  } catch (e) {}

  for (var c = 1; c <= lastCol; c++) {
    var h = header[c - 1] || guessHeaderFromSheet_v620_(sheet, c);
    try {
      var width = fastWidthByHeader_v620_(h, isLarge);
      if (!isLarge && sheet.getColumnWidth(c) < width.min) sheet.setColumnWidth(c, width.min);
      if (sheet.getColumnWidth(c) > width.max) sheet.setColumnWidth(c, width.max);
      if (isLarge && width.fixed) sheet.setColumnWidth(c, width.fixed);

      if (/상품명|메모|비고|사유|주문번호|대표검색필터명|필터명/.test(h)) {
        sheet.getRange(1, c, Math.min(lastRow, 500), 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      }
      if (/년|월|일/.test(h) && lastRow > 1) {
        sheet.getRange(2, c, lastRow - 1, 1).setNumberFormat('0');
      }
      if (/주문번호|상품번호|계정ID/.test(h) && lastRow > 1) {
        sheet.getRange(2, c, lastRow - 1, 1).setNumberFormat('@');
      }
      if (/순수매출액|실제정산금액|예상정산금액|정산기준금액|매입금액|예상이익|총매출|취소|수수료|공급|부가세|정산|매출액|금액/.test(h) && lastRow > 1) {
        sheet.getRange(2, c, lastRow - 1, 1).setNumberFormat('₩#,##0');
      }
      if (/율/.test(h) && lastRow > 1) {
        sheet.getRange(2, c, lastRow - 1, 1).setNumberFormat('0.00%');
      }
      if (/수집수|매출상품|주문|미정산건수|30일초과건수|미판매수|수량|건수|고객수/.test(h) && !/주문번호/.test(h) && lastRow > 1) {
        sheet.getRange(2, c, lastRow - 1, 1).setNumberFormat('#,##0');
      }
    } catch (e) {}
  }

  // 대시보드의 중간 섹션 헤더만 가볍게 보정합니다.
  if (sheet.getName() === CONFIG.SHEETS.DASHBOARD) {
    styleDashboardSectionHeadersFast_v620_(sheet, Math.min(lastRow, 140), lastCol);
  }

  return { formatted: true, sheet: sheet.getName(), large: isLarge };
}

function styleDashboardSectionHeadersFast_v620_(sheet, rowLimit, lastCol) {
  try {
    var values = sheet.getRange(1, 1, rowLimit, Math.min(lastCol, 16)).getValues();
    for (var r = 0; r < values.length; r++) {
      if (r === 0 || String(values[r][0] || '').trim() === '구분') {
        sheet.getRange(r + 1, 1, 1, lastCol)
          .setBackground('#d9eaf7')
          .setFontWeight('bold')
          .setHorizontalAlignment('center');
      }
    }
  } catch (e) {}
}

function fastWidthByHeader_v620_(header, isLarge) {
  var h = String(header || '').trim();
  var result = { min: 55, max: 150, fixed: 0 };
  if (/상품명/.test(h)) result = { min: 180, max: 360, fixed: isLarge ? 320 : 0 };
  else if (/메모|비고|사유/.test(h)) result = { min: 160, max: 260, fixed: isLarge ? 220 : 0 };
  else if (/대표검색필터명/.test(h)) result = { min: 150, max: 210, fixed: isLarge ? 190 : 0 };
  else if (/주문번호/.test(h)) result = { min: 120, max: 180, fixed: isLarge ? 160 : 0 };
  else if (/고객명/.test(h)) result = { min: 90, max: 120, fixed: isLarge ? 110 : 0 };
  else if (/브랜드명/.test(h)) result = { min: 100, max: 140, fixed: isLarge ? 130 : 0 };
  else if (/상품번호/.test(h)) result = { min: 110, max: 130, fixed: isLarge ? 120 : 0 };
  else if (/계정ID/.test(h)) result = { min: 95, max: 120, fixed: isLarge ? 110 : 0 };
  else if (/년|월|일/.test(h)) result = { min: 45, max: 70, fixed: isLarge ? 60 : 0 };
  else if (/율/.test(h)) result = { min: 75, max: 95, fixed: isLarge ? 85 : 0 };
  else if (/금액|매출|정산|매입|이익|수수료|부가세|공급/.test(h)) result = { min: 105, max: 135, fixed: isLarge ? 125 : 0 };
  return result;
}

function guessHeaderFromSheet_v620_(sheet, col) {
  try {
    var rowLimit = Math.min(sheet.getLastRow(), 80);
    var vals = sheet.getRange(1, col, rowLimit, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      var text = String(vals[i][0] || '').trim();
      if (/상품명|메모|비고|주문번호|브랜드명|금액|율|계정ID|필터명|정산|매입|이익/.test(text)) return text;
    }
  } catch (e) {}
  return '';
}
