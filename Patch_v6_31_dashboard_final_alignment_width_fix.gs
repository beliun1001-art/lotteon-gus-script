/**
 * LOTTEON v6.31 final dashboard alignment and width fix
 *
 * 문제:
 * - 대시보드 요약 영역 4~29행의 값 숫자가 좌측 정렬로 남았습니다.
 * - 브랜드TOP 영역 32~91행의 구분열(A열)이 좁게 보여 자동 너비 조정이 체감되지 않았습니다.
 *
 * 수정:
 * - 대시보드 A열 구분열 폭을 92px로 강제합니다.
 * - 요약 영역 C4:C29 값 영역을 우측 정렬로 강제합니다.
 * - 표 영역은 헤더명 기준으로 숫자/금액/비율/건수 컬럼을 우측 정렬로 강제합니다.
 * - 텍스트 컬럼은 A/B/메모/다음작업/최근수집일만 좌측 또는 가운데 정렬합니다.
 * - v6.30 대시보드 생성/표시서식 정리/열너비 자동조정 후 마지막 단계에서 다시 한 번 적용됩니다.
 */

var LOTTEON_PATCH_V631_DASHBOARD_FINAL_ALIGNMENT_WIDTH_FIX_LOADED = true;

var __baseHardFixDashboardDisplay_v631 = typeof hardFixDashboardDisplay_v630_ === 'function' ? hardFixDashboardDisplay_v630_ : null;
var __baseApplyDisplayStandardsOnlyFast_v631 = typeof applyDisplayStandardsOnlyFast_v623 === 'function' ? applyDisplayStandardsOnlyFast_v623 : null;
var __baseApplyFastOutputSheetFormatting_v631 = typeof applyFastOutputSheetFormatting_v620_ === 'function' ? applyFastOutputSheetFormatting_v620_ : null;
var __baseRunColumnWidthStep_v631 = typeof runColumnWidthAutoAdjustStep_v623 === 'function' ? runColumnWidthAutoAdjustStep_v623 : null;
var __baseRefreshDashboardFastOnly_v631 = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;

hardFixDashboardDisplay_v630_ = function() {
  var result = __baseHardFixDashboardDisplay_v631 ? __baseHardFixDashboardDisplay_v631.apply(this, arguments) : null;
  forceDashboardFinalAlignmentWidth_v631_();
  return result || { ok: true };
};

applyDisplayStandardsOnlyFast_v623 = function() {
  var result = __baseApplyDisplayStandardsOnlyFast_v631 ? __baseApplyDisplayStandardsOnlyFast_v631.apply(this, arguments) : null;
  forceDashboardFinalAlignmentWidth_v631_();
  try { SpreadsheetApp.getUi().alert('표시서식 정리 완료\n\n대시보드 요약 숫자 우측정렬과 구분열 폭을 최종 보정했습니다.'); } catch (e) {}
  return result || { ok: true };
};

applyFastOutputSheetFormatting_v620_ = function() {
  var result = __baseApplyFastOutputSheetFormatting_v631 ? __baseApplyFastOutputSheetFormatting_v631.apply(this, arguments) : null;
  forceDashboardFinalAlignmentWidth_v631_();
  return result || { ok: true };
};

runColumnWidthAutoAdjustStep_v623 = function() {
  var result = __baseRunColumnWidthStep_v631 ? __baseRunColumnWidthStep_v631.apply(this, arguments) : null;
  forceDashboardFinalAlignmentWidth_v631_();
  return result;
};

refreshDashboardFastOnly = function() {
  var result = __baseRefreshDashboardFastOnly_v631 ? __baseRefreshDashboardFastOnly_v631.apply(this, arguments) : null;
  forceDashboardFinalAlignmentWidth_v631_();
  return result;
};

function forceDashboardFinalAlignmentWidth_v631_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!sheet || sheet.getLastRow() < 1) return { skipped: true };

  var lastRow = Math.min(sheet.getLastRow(), 260);
  var lastCol = Math.min(sheet.getLastColumn(), 16);
  if (lastCol < 1) return { skipped: true };

  try {
    // 구분열은 브랜드TOP/상품갈이가 잘 보이도록 넓게 고정합니다.
    sheet.setColumnWidth(1, 92);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 92);
    sheet.setColumnWidth(4, 88);
    sheet.setColumnWidth(5, 78);
    sheet.setColumnWidth(6, 118);
    sheet.setColumnWidth(7, 120);
    sheet.setColumnWidth(8, 138);
    sheet.setColumnWidth(9, 120);
    sheet.setColumnWidth(10, 118);
    sheet.setColumnWidth(11, 114);
    sheet.setColumnWidth(12, 90);
    sheet.setColumnWidth(13, 94);
    sheet.setColumnWidth(14, 94);
    sheet.setColumnWidth(15, 136);
    sheet.setColumnWidth(16, 230);
  } catch (e) {}

  try {
    sheet.setRowHeights(1, lastRow, 22);
  } catch (e) {}

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headerRows = [];
  for (var r = 0; r < values.length; r++) {
    var first = String(values[r][0] || '').replace(/\n/g, '').trim();
    if (r === 0 || first === '구분') headerRows.push(r + 1);
  }

  // 전체 기본은 우측 정렬. 텍스트 컬럼만 뒤에서 좌측/중앙으로 재지정합니다.
  if (lastRow > 1) {
    try {
      sheet.getRange(2, 1, lastRow - 1, lastCol)
        .setHorizontalAlignment('right')
        .setVerticalAlignment('middle')
        .setFontWeight('normal')
        .setBackground(null)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    } catch (e) {}
  }

  // 요약 영역: C4:C29는 숫자/금액/비율/건수 값 영역이므로 우측 고정합니다.
  if (lastRow >= 4) {
    var summaryEnd = Math.min(29, lastRow);
    try { sheet.getRange(4, 3, summaryEnd - 3, 1).setHorizontalAlignment('right'); } catch (e) {}
  }

  // 요약 텍스트 컬럼
  try { if (lastRow > 1) sheet.getRange(2, 1, Math.min(28, lastRow - 1), 2).setHorizontalAlignment('left'); } catch (e) {}
  try { if (lastRow > 1 && lastCol >= 16) sheet.getRange(2, 16, Math.min(28, lastRow - 1), 1).setHorizontalAlignment('left'); } catch (e) {}

  // 섹션별 헤더/데이터 포맷
  headerRows.forEach(function(rowNo, idx) {
    var nextHeader = headerRows[idx + 1] || (lastRow + 1);
    var endRow = Math.min(nextHeader - 1, lastRow);
    try {
      sheet.getRange(rowNo, 1, 1, lastCol)
        .setBackground('#d9eaf7')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      sheet.setRowHeight(rowNo, 34);
    } catch (e) {}
    if (endRow <= rowNo) return;

    var rowCount = endRow - rowNo;
    for (var c = 1; c <= lastCol; c++) {
      var header = normalizeDashboardFinalHeader_v631_(values[rowNo - 1][c - 1]);
      if (!header) continue;
      try {
        var rg = sheet.getRange(rowNo + 1, c, rowCount, 1);
        if (isDashboardTextColumn_v631_(header, c)) {
          if (/최근수집일|날짜|일자/.test(header)) rg.setNumberFormat('@').setHorizontalAlignment('center');
          else rg.setNumberFormat('@').setHorizontalAlignment('left');
        } else if (/율/.test(header)) {
          rg.setNumberFormat('0.00%').setHorizontalAlignment('right');
        } else if (/금액|매출|정산|매입|이익|수수료|부가세|공급/.test(header) && !/상품수|건수|수량/.test(header)) {
          rg.setNumberFormat('#,##0').setHorizontalAlignment('right');
        } else if (/수집수|매출상품|주문|건수|수량|미판매수|30일초과/.test(header) && !/주문번호/.test(header)) {
          rg.setNumberFormat('#,##0').setHorizontalAlignment('right');
        } else {
          // 헤더 판정이 애매해도 표 데이터의 숫자형 컬럼은 우측 유지.
          rg.setHorizontalAlignment('right');
        }
      } catch (e) {}
    }
  });

  // 구분/브랜드명/메모 컬럼은 최종적으로 좌측 고정합니다.
  [1, 2, 15, 16].forEach(function(col) {
    if (col <= lastCol && lastRow > 1) {
      try { sheet.getRange(2, col, lastRow - 1, 1).setHorizontalAlignment('left'); } catch (e) {}
    }
  });

  try { log_('patch_v631_dashboard_final_alignment_width', 'rows=' + lastRow + ' cols=' + lastCol); } catch (e) {}
  return { ok: true };
}

function normalizeDashboardFinalHeader_v631_(v) {
  return String(v == null ? '' : v).replace(/\n/g, '').replace(/\s+/g, '').trim();
}

function isDashboardTextColumn_v631_(header, col) {
  header = normalizeDashboardFinalHeader_v631_(header);
  if (col === 1 || col === 2) return true;
  if (/메모|다음작업|최근수집일|날짜|일자|구분|브랜드명|항목/.test(header)) return true;
  return false;
}
