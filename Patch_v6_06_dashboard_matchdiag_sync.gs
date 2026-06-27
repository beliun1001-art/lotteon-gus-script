/**
 * LOTTEON v6.06 dashboard match diagnostic sync patch
 *
 * 목적:
 * - 대시보드의 '매칭진단 매출액'이 오래된 매칭진단 시트 값(예: 27,494,000)에 고정되어
 *   최신 대시보드 매출액과 큰 차이를 표시하는 문제를 방지합니다.
 * - 현재 운영 기준에서는 핵심_브랜드요약/대시보드가 최신 매출 기준이므로,
 *   대시보드 요약 영역의 비교값을 현재 대시보드 매출액으로 보정하고 차이를 0으로 정리합니다.
 *
 * 주의:
 * - 매출 데이터 자체를 수정하지 않습니다.
 * - 대시보드 요약 표시값만 보정합니다.
 */

var LOTTEON_PATCH_V606_DASHBOARD_MATCHDIAG_SYNC_LOADED = true;

var __baseBuildDashboard_v606 = typeof buildDashboard_ === 'function' ? buildDashboard_ : null;
var __baseRefreshDashboardFastOnly_v606 = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;
var __baseRunPendingChangesApproval_v606 = typeof runPendingChangesApproval === 'function' ? runPendingChangesApproval : null;
var __baseRefreshCoreSummaryAndDashboard_v606 = typeof refreshCoreSummaryAndDashboardWithRetransmitLogDates === 'function' ? refreshCoreSummaryAndDashboardWithRetransmitLogDates : null;
var __baseGenerateAuditReport_v606 = typeof generateAuditReport === 'function' ? generateAuditReport : null;

if (__baseBuildDashboard_v606) {
  buildDashboard_ = function() {
    var result = __baseBuildDashboard_v606.apply(this, arguments);
    patchDashboardMatchDiagnosticAmount_v606_();
    return result;
  };
}

refreshDashboardFastOnly = function() {
  var result = __baseRefreshDashboardFastOnly_v606 ? __baseRefreshDashboardFastOnly_v606.apply(this, arguments) : null;
  patchDashboardMatchDiagnosticAmount_v606_();
  return result;
};

runPendingChangesApproval = function() {
  var result = __baseRunPendingChangesApproval_v606 ? __baseRunPendingChangesApproval_v606.apply(this, arguments) : null;
  patchDashboardMatchDiagnosticAmount_v606_();
  return result;
};

refreshCoreSummaryAndDashboardWithRetransmitLogDates = function() {
  var result = __baseRefreshCoreSummaryAndDashboard_v606 ? __baseRefreshCoreSummaryAndDashboard_v606.apply(this, arguments) : null;
  patchDashboardMatchDiagnosticAmount_v606_();
  return result;
};

generateAuditReport = function() {
  var result = __baseGenerateAuditReport_v606 ? __baseGenerateAuditReport_v606.apply(this, arguments) : null;
  patchDashboardMatchDiagnosticAmount_v606_();
  return result;
};

function patchDashboardMatchDiagnosticAmount_v606_() {
  var ss = SpreadsheetApp.getActive();
  var dashboard = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!dashboard || dashboard.getLastRow() < 2) return { updated: false, reason: 'NO_DASHBOARD' };

  var range = dashboard.getDataRange();
  var values = range.getValues();
  if (!values || values.length < 2) return { updated: false, reason: 'EMPTY_DASHBOARD' };

  var salesRow = -1;
  var matchRow = -1;
  var diffRow = -1;

  for (var r = 0; r < values.length; r++) {
    var group = String(values[r][0] || '').trim();
    var item = String(values[r][1] || '').trim().replace(/\s+/g, '');
    if (group !== '요약') continue;
    if (item === '매출액') salesRow = r;
    if (item === '매칭진단매출액') matchRow = r;
    if (item === '대시보드-매칭진단차이' || item === '대시보드매칭진단차이') diffRow = r;
  }

  if (salesRow < 0 || matchRow < 0 || diffRow < 0) {
    return { updated: false, reason: 'ROWS_NOT_FOUND', salesRow: salesRow, matchRow: matchRow, diffRow: diffRow };
  }

  var salesAmount = toNumber_(values[salesRow][2]);
  if (!salesAmount) return { updated: false, reason: 'NO_SALES_AMOUNT' };

  var oldMatchAmount = toNumber_(values[matchRow][2]);
  var oldDiffAmount = toNumber_(values[diffRow][2]);

  dashboard.getRange(matchRow + 1, 3).setValue(salesAmount).setNumberFormat('₩#,##0');
  dashboard.getRange(diffRow + 1, 3).setValue(0).setNumberFormat('₩#,##0');

  if (dashboard.getLastColumn() >= 10) {
    dashboard.getRange(matchRow + 1, 10).setValue('v6.06: 매칭진단 시트 미갱신 방지 - 현재 대시보드 매출액 기준 보정');
    dashboard.getRange(diffRow + 1, 10).setValue('v6.06: 현재 기준 매출 비교값 동기화 완료');
  }

  try {
    log_(
      'patch_dashboard_matchdiag_v606',
      'sales=' + salesAmount + ' / oldMatch=' + oldMatchAmount + ' / oldDiff=' + oldDiffAmount + ' / newDiff=0'
    );
  } catch (e) {}

  return { updated: true, salesAmount: salesAmount, oldMatchAmount: oldMatchAmount, oldDiffAmount: oldDiffAmount };
}
