/**
 * LOTTEON v6.15 dashboard dynamic header styling patch
 *
 * 목적:
 * - 대시보드 데이터 행 수가 바뀌어도 섹션 헤더 행을 자동 감지해 색상/굵은 글씨를 적용합니다.
 * - 기존 데이터가 줄어든 뒤 남는 과거 헤더 서식도 자동으로 정리합니다.
 *
 * 적용 기준:
 * - 1행: 전체 헤더
 * - A열 값이 '구분'인 행: 섹션 헤더
 *   예: 브랜드TOP 헤더, 상품갈이 헤더, 조치필요 헤더
 */

var LOTTEON_PATCH_V615_DASHBOARD_HEADER_STYLE_AUTO_LOADED = true;

var __baseBuildDashboard_v615 = typeof buildDashboard_ === 'function' ? buildDashboard_ : null;
var __baseRefreshDashboardFastOnly_v615 = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;
var __baseRunPendingChangesApproval_v615 = typeof runPendingChangesApproval === 'function' ? runPendingChangesApproval : null;
var __baseRefreshCoreSummaryAndDashboard_v615 = typeof refreshCoreSummaryAndDashboardWithRetransmitLogDates === 'function' ? refreshCoreSummaryAndDashboardWithRetransmitLogDates : null;
var __baseRebuildDashboardBrandBased_v615 = typeof rebuildDashboardBrandBased_v611_ === 'function' ? rebuildDashboardBrandBased_v611_ : null;

if (__baseBuildDashboard_v615) {
  buildDashboard_ = function() {
    var result = __baseBuildDashboard_v615.apply(this, arguments);
    applyDashboardDynamicHeaderStyle_v615_();
    return result;
  };
}

refreshDashboardFastOnly = function() {
  var result = __baseRefreshDashboardFastOnly_v615 ? __baseRefreshDashboardFastOnly_v615.apply(this, arguments) : null;
  applyDashboardDynamicHeaderStyle_v615_();
  return result;
};

runPendingChangesApproval = function() {
  var result = __baseRunPendingChangesApproval_v615 ? __baseRunPendingChangesApproval_v615.apply(this, arguments) : null;
  applyDashboardDynamicHeaderStyle_v615_();
  return result;
};

refreshCoreSummaryAndDashboardWithRetransmitLogDates = function() {
  var result = __baseRefreshCoreSummaryAndDashboard_v615 ? __baseRefreshCoreSummaryAndDashboard_v615.apply(this, arguments) : null;
  applyDashboardDynamicHeaderStyle_v615_();
  return result;
};

if (__baseRebuildDashboardBrandBased_v615) {
  rebuildDashboardBrandBased_v611_ = function() {
    var result = __baseRebuildDashboardBrandBased_v615.apply(this, arguments);
    applyDashboardDynamicHeaderStyle_v615_();
    return result;
  };
}

function applyDashboardDynamicHeaderStyle_v615_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!sheet) return { updated: 0, reason: 'NO_DASHBOARD' };

  var lastRow = Math.max(sheet.getLastRow(), 1);
  var lastCol = Math.max(sheet.getLastColumn(), 11);
  var maxRowsToClean = Math.min(sheet.getMaxRows(), Math.max(lastRow + 30, 120));
  var maxColsToClean = Math.min(sheet.getMaxColumns(), Math.max(lastCol, 11));

  // 과거에 남은 헤더/색상 서식 정리. 값/번호서식은 건드리지 않습니다.
  var cleanRange = sheet.getRange(1, 1, maxRowsToClean, maxColsToClean);
  cleanRange
    .setBackground(null)
    .setFontWeight('normal')
    .setFontColor(null)
    .setVerticalAlignment('middle');

  if (lastRow < 1) return { updated: 0, reason: 'EMPTY_DASHBOARD' };

  var values = sheet.getRange(1, 1, lastRow, maxColsToClean).getValues();
  var headerRows = [];
  for (var r = 0; r < values.length; r++) {
    if (isDashboardHeaderRow_v615_(values[r], r)) headerRows.push(r + 1);
  }

  var sectionHeaderColor = '#d9eaf7';
  var topHeaderColor = '#cfe2f3';
  var headerFontColor = '#000000';

  headerRows.forEach(function(rowNo) {
    var bg = rowNo === 1 ? topHeaderColor : sectionHeaderColor;
    sheet.getRange(rowNo, 1, 1, maxColsToClean)
      .setBackground(bg)
      .setFontWeight('bold')
      .setFontColor(headerFontColor)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  });

  // 데이터 행은 왼쪽 정렬/기본 배경. 금액/숫자 서식은 기존 v6.13/v6.14 로직 유지.
  try {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, lastRow, maxColsToClean).setFontFamily('Arial').setFontSize(10);
    sheet.autoResizeColumns(1, Math.min(maxColsToClean, 11));
  } catch (e) {}

  installDashboardHeaderConditionalRules_v615_(sheet, maxRowsToClean, maxColsToClean);

  try {
    log_('patch_dashboard_header_style_v615', 'headerRows=' + headerRows.join(',') + ' / cleanedRows=' + maxRowsToClean);
  } catch (e) {}

  return { updated: headerRows.length, headerRows: headerRows };
}

function isDashboardHeaderRow_v615_(row, zeroBasedRowIndex) {
  var a = String(row[0] || '').trim();
  var b = String(row[1] || '').trim();
  var c = String(row[2] || '').trim();

  if (zeroBasedRowIndex === 0) return true;
  if (a === '구분' && (b === '브랜드명' || b === '항목' || c === '대표검색필터명')) return true;
  if (a === '구분' && /브랜드명|항목/.test(b + ' ' + c)) return true;
  return false;
}

function installDashboardHeaderConditionalRules_v615_(sheet, maxRows, maxCols) {
  try {
    var targetRange = sheet.getRange(1, 1, maxRows, maxCols);
    var rules = sheet.getConditionalFormatRules() || [];
    var kept = rules.filter(function(rule) {
      var bc = rule.getBooleanCondition && rule.getBooleanCondition();
      if (!bc) return true;
      var vals = bc.getCriteriaValues && bc.getCriteriaValues();
      var formula = vals && vals.length ? String(vals[0] || '') : '';
      return formula.indexOf('LOTTEON_DASHBOARD_HEADER_V615') < 0;
    });

    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=OR(ROW()=1,$A1="구분")+N("LOTTEON_DASHBOARD_HEADER_V615")')
      .setBackground('#d9eaf7')
      .setBold(true)
      .setRanges([targetRange])
      .build();

    kept.push(rule);
    sheet.setConditionalFormatRules(kept);
  } catch (e) {
    // 조건부서식 API가 실패해도 직접 서식은 이미 적용되어 있으므로 무시합니다.
  }
}
