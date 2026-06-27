/**
 * LOTTEON v6.07 dashboard total sent count sync patch
 *
 * 목적:
 * - 대시보드 요약의 '쿠팡전송상품수'가 핵심_브랜드요약 전송수 합계 또는 과거 수동입력/중복 행 기준으로 표시되는 문제를 방지합니다.
 * - 운영 기준 전송수는 필터별_상품수.API_totalCount 총합입니다.
 * - 대시보드 요약의 쿠팡전송상품수와 전체 매출상품률을 필터별_상품수 기준으로 직접 보정합니다.
 *
 * 주의:
 * - 필터별_상품수 원본은 수정하지 않습니다.
 * - 대시보드 요약 표시값만 보정합니다.
 */

var LOTTEON_PATCH_V607_DASHBOARD_TOTAL_SENT_SYNC_LOADED = true;

var __baseBuildDashboard_v607 = typeof buildDashboard_ === 'function' ? buildDashboard_ : null;
var __baseRefreshDashboardFastOnly_v607 = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;
var __baseRunPendingChangesApproval_v607 = typeof runPendingChangesApproval === 'function' ? runPendingChangesApproval : null;
var __baseRefreshCoreSummaryAndDashboard_v607 = typeof refreshCoreSummaryAndDashboardWithRetransmitLogDates === 'function' ? refreshCoreSummaryAndDashboardWithRetransmitLogDates : null;

if (__baseBuildDashboard_v607) {
  buildDashboard_ = function() {
    var result = __baseBuildDashboard_v607.apply(this, arguments);
    patchDashboardTotalSentCountFromFilterApi_v607_();
    return result;
  };
}

refreshDashboardFastOnly = function() {
  var result = __baseRefreshDashboardFastOnly_v607 ? __baseRefreshDashboardFastOnly_v607.apply(this, arguments) : null;
  patchDashboardTotalSentCountFromFilterApi_v607_();
  return result;
};

runPendingChangesApproval = function() {
  var result = __baseRunPendingChangesApproval_v607 ? __baseRunPendingChangesApproval_v607.apply(this, arguments) : null;
  patchDashboardTotalSentCountFromFilterApi_v607_();
  return result;
};

refreshCoreSummaryAndDashboardWithRetransmitLogDates = function() {
  var result = __baseRefreshCoreSummaryAndDashboard_v607 ? __baseRefreshCoreSummaryAndDashboard_v607.apply(this, arguments) : null;
  patchDashboardTotalSentCountFromFilterApi_v607_();
  return result;
};

function patchDashboardTotalSentCountFromFilterApi_v607_() {
  var ss = SpreadsheetApp.getActive();
  var dashboard = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!dashboard || dashboard.getLastRow() < 2) return { updated: false, reason: 'NO_DASHBOARD' };

  var totalSent = sumFilterApiTotalCount_v607_();
  if (!totalSent) return { updated: false, reason: 'NO_FILTER_TOTAL' };

  var range = dashboard.getDataRange();
  var values = range.getValues();
  if (!values || values.length < 2) return { updated: false, reason: 'EMPTY_DASHBOARD' };

  var sentRow = -1;
  var salesProductRow = -1;
  var totalRateRow = -1;

  for (var r = 0; r < values.length; r++) {
    var group = String(values[r][0] || '').trim();
    var item = String(values[r][1] || '').trim().replace(/\s+/g, '');
    if (group !== '요약') continue;
    if (item === '쿠팡전송상품수') sentRow = r;
    if (item === '매출상품수') salesProductRow = r;
    if (item === '전체매출상품률') totalRateRow = r;
  }

  if (sentRow < 0) return { updated: false, reason: 'SENT_ROW_NOT_FOUND' };

  var oldSent = toNumber_(values[sentRow][2]);
  dashboard.getRange(sentRow + 1, 3).setValue(totalSent).setNumberFormat('#,##0');
  if (dashboard.getLastColumn() >= 10) {
    dashboard.getRange(sentRow + 1, 10).setValue('v6.07: 필터별_상품수 API_totalCount 총합 기준');
  }

  var salesProduct = salesProductRow >= 0 ? toNumber_(values[salesProductRow][2]) : 0;
  if (totalRateRow >= 0 && totalSent > 0) {
    dashboard.getRange(totalRateRow + 1, 3).setValue(salesProduct / totalSent).setNumberFormat('0.00%');
    if (dashboard.getLastColumn() >= 10) {
      dashboard.getRange(totalRateRow + 1, 10).setValue('v6.07: 매출상품수 ÷ 필터별_상품수 API_totalCount 총합');
    }
  }

  try {
    log_('patch_dashboard_total_sent_v607', 'oldSent=' + oldSent + ' / newSent=' + totalSent + ' / salesProduct=' + salesProduct);
  } catch (e) {}

  return { updated: true, oldSent: oldSent, newSent: totalSent, salesProduct: salesProduct };
}

function sumFilterApiTotalCount_v607_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.FILTERS);
  if (!sheet || sheet.getLastRow() < 2) return 0;

  var values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return 0;

  var header = values[0];
  var colTotal = findHeaderIndex_v607_(header, ['API_totalCount', 'APItotalCount']);
  if (colTotal < 0) return 0;

  var total = 0;
  for (var r = 1; r < values.length; r++) {
    total += toNumber_(values[r][colTotal]);
  }
  return total;
}

function findHeaderIndex_v607_(headerRow, candidates) {
  var normalized = (headerRow || []).map(function(h) { return normalizeHeaderKey_v607_(h); });
  for (var i = 0; i < candidates.length; i++) {
    var key = normalizeHeaderKey_v607_(candidates[i]);
    var idx = normalized.indexOf(key);
    if (idx >= 0) return idx;
  }
  return -1;
}

function normalizeHeaderKey_v607_(v) {
  return String(v || '').replace(/\s+/g, '').replace(/\n/g, '').replace(/_/g, '').trim();
}
