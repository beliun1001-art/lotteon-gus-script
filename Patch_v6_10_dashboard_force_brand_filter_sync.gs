/**
 * LOTTEON v6.10 force dashboard brand/filter sync patch
 *
 * 목적:
 * - v6.09 이후에도 대시보드 확대TOP/상품갈이/조치필요의 브랜드명, 대표검색필터명, 메모가
 *   과거 값으로 남는 문제를 강제로 보정합니다.
 * - 기존 대표검색필터명이 필터별_상품수에 존재해도, 같은 브랜드/계정/전송수의 최신 관리필터가 있으면
 *   관리필터를 우선 사용합니다.
 *
 * 브랜드 정규화:
 * - 앞 숫자 또는 숫자_ 제거: 1K2, 1_K2 -> K2
 * - 뒤 _번호 제거: 플라스틱아일랜드_01 -> 플라스틱아일랜드
 */

var LOTTEON_PATCH_V610_DASHBOARD_FORCE_BRAND_FILTER_SYNC_LOADED = true;

var __baseBuildDashboard_v610 = typeof buildDashboard_ === 'function' ? buildDashboard_ : null;
var __baseRefreshDashboardFastOnly_v610 = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;
var __baseRunPendingChangesApproval_v610 = typeof runPendingChangesApproval === 'function' ? runPendingChangesApproval : null;
var __baseRefreshCoreSummaryAndDashboard_v610 = typeof refreshCoreSummaryAndDashboardWithRetransmitLogDates === 'function' ? refreshCoreSummaryAndDashboardWithRetransmitLogDates : null;

if (__baseBuildDashboard_v610) {
  buildDashboard_ = function() {
    var result = __baseBuildDashboard_v610.apply(this, arguments);
    forcePatchDashboardBrandFilter_v610_();
    SpreadsheetApp.flush();
    return result;
  };
}

refreshDashboardFastOnly = function() {
  var result = __baseRefreshDashboardFastOnly_v610 ? __baseRefreshDashboardFastOnly_v610.apply(this, arguments) : null;
  forcePatchDashboardBrandFilter_v610_();
  SpreadsheetApp.flush();
  return result;
};

runPendingChangesApproval = function() {
  var result = __baseRunPendingChangesApproval_v610 ? __baseRunPendingChangesApproval_v610.apply(this, arguments) : null;
  forcePatchDashboardBrandFilter_v610_();
  SpreadsheetApp.flush();
  return result;
};

refreshCoreSummaryAndDashboardWithRetransmitLogDates = function() {
  var result = __baseRefreshCoreSummaryAndDashboard_v610 ? __baseRefreshCoreSummaryAndDashboard_v610.apply(this, arguments) : null;
  forcePatchDashboardBrandFilter_v610_();
  SpreadsheetApp.flush();
  return result;
};

function forcePatchDashboardBrandFilter_v610_() {
  var ss = SpreadsheetApp.getActive();
  var dashboard = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!dashboard || dashboard.getLastRow() < 2) return { updated: 0, reason: 'NO_DASHBOARD' };

  var lookup = buildFilterLookupForDashboard_v610_();
  if (!lookup.rows.length) return { updated: 0, reason: 'NO_FILTER_LOOKUP' };

  var lastRow = dashboard.getLastRow();
  var lastCol = dashboard.getLastColumn();
  var values = dashboard.getRange(1, 1, lastRow, lastCol).getValues();
  var updated = 0;
  var targetGroups = { '확대TOP': true, '상품갈이': true, '조치필요': true };

  for (var r = 1; r < values.length; r++) {
    var group = String(values[r][0] || '').trim();
    if (!targetGroups[group]) continue;

    var oldBrand = String(values[r][1] || '').trim();
    var oldFilter = String(values[r][2] || '').trim();
    var oldAccount = String(values[r][3] || '').trim();
    var oldTotal = toNumber_(values[r][4]);
    var salesProduct = toNumber_(values[r][5]);

    var canonical = canonicalBrandForDashboard_v610_(oldBrand) || canonicalBrandFromFilterForDashboard_v610_(oldFilter);
    var item = chooseBestFilterForDashboard_v610_(lookup, canonical, oldFilter, oldAccount, oldTotal);
    if (!item) continue;

    var newRow = values[r].slice();
    newRow[1] = item.brand;
    newRow[2] = item.filterName;
    newRow[3] = item.accountId || oldAccount;
    newRow[4] = item.total;
    if (item.total > 0) newRow[6] = salesProduct / item.total;

    if (newRow.length > 10) newRow[10] = rewriteDashboardMemo_v610_(String(newRow[10] || ''), item.total);

    var changed = false;
    for (var c = 1; c <= Math.min(10, newRow.length - 1); c++) {
      if (String(values[r][c]) !== String(newRow[c])) {
        changed = true;
        break;
      }
    }

    if (changed) {
      dashboard.getRange(r + 1, 1, 1, lastCol).setValues([newRow]);
      updated++;
    }
  }

  try {
    dashboard.getRange(1, 5, lastRow, 1).setNumberFormat('#,##0');
    dashboard.getRange(1, 7, lastRow, 1).setNumberFormat('0.00%');
    log_('patch_dashboard_brand_filter_v610', 'updated=' + updated + ' / force=Y');
  } catch (e) {}

  return { updated: updated };
}

function buildFilterLookupForDashboard_v610_() {
  var result = { rows: [], byBrand: {} };
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.FILTERS);
  if (!sheet || sheet.getLastRow() < 2) return result;

  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var colFilter = findHeaderIndex_v610_(header, ['검색필터명']);
  var colTotal = findHeaderIndex_v610_(header, ['API_totalCount', 'APItotalCount']);
  var colAccount = findHeaderIndex_v610_(header, ['쿠팡계정ID']);
  if (colFilter < 0 || colTotal < 0) return result;

  for (var r = 1; r < values.length; r++) {
    var filterName = String(values[r][colFilter] || '').trim();
    if (!filterName || filterName.indexOf(CONFIG.FILTER_PREFIX) !== 0) continue;

    var brand = canonicalBrandFromFilterForDashboard_v610_(filterName);
    var key = normalizeBrandKeyForDashboard_v610_(brand);
    var accountId = colAccount >= 0 ? String(values[r][colAccount] || '').trim() : accountIdFromFilterForDashboard_v610_(filterName);
    var total = toNumber_(values[r][colTotal]);
    var item = {
      filterName: filterName,
      brand: brand,
      key: key,
      accountId: accountId,
      total: total,
      managedScore: managedFilterScore_v610_(filterName)
    };

    result.rows.push(item);
    if (key) {
      if (!result.byBrand[key]) result.byBrand[key] = [];
      result.byBrand[key].push(item);
    }
  }
  return result;
}

function chooseBestFilterForDashboard_v610_(lookup, canonicalBrand, oldFilter, oldAccount, oldTotal) {
  var key = normalizeBrandKeyForDashboard_v610_(canonicalBrand);
  var list = key ? (lookup.byBrand[key] || []) : [];
  if (!list.length) return null;

  var scored = list.map(function(item) {
    var score = 0;
    if (oldTotal && item.total === oldTotal) score += 1000;
    if (oldAccount && item.accountId === oldAccount) score += 300;
    score += item.managedScore || 0;
    if (item.filterName !== oldFilter) score += 20;
    if (item.filterName === oldFilter) score -= 10;
    score += Math.min(item.total || 0, 999) / 10000;
    return { item: item, score: score };
  });

  scored.sort(function(a, b) {
    if (b.score !== a.score) return b.score - a.score;
    return String(a.item.filterName || '').localeCompare(String(b.item.filterName || ''));
  });

  return scored[0] ? scored[0].item : null;
}

function managedFilterScore_v610_(filterName) {
  var tail = filterNameTailAfterSecondUnderscore_v610_(filterName);
  var score = 0;
  if (/^\d+_?/.test(tail)) score += 100;
  if (/_\d+$/.test(tail)) score += 50;
  return score;
}

function canonicalBrandFromFilterForDashboard_v610_(filterName) {
  return canonicalBrandForDashboard_v610_(filterNameTailAfterSecondUnderscore_v610_(filterName));
}

function filterNameTailAfterSecondUnderscore_v610_(filterName) {
  var text = String(filterName || '').trim();
  var first = text.indexOf('_');
  if (first < 0) return text;
  var second = text.indexOf('_', first + 1);
  if (second < 0) return text;
  return text.slice(second + 1).trim();
}

function canonicalBrandForDashboard_v610_(brandName) {
  return String(brandName || '')
    .trim()
    .replace(/^\d+_?/, '')
    .replace(/_\d+$/, '')
    .trim();
}

function normalizeBrandKeyForDashboard_v610_(brandName) {
  return canonicalBrandForDashboard_v610_(brandName).toLowerCase().replace(/\s+/g, '').replace(/[\[\]\(\)\{\}\-_]/g, '').trim();
}

function accountIdFromFilterForDashboard_v610_(filterName) {
  var text = String(filterName || '').trim();
  if (text.indexOf('롯백_01_') === 0) return 'beliun1021';
  if (text.indexOf('롯백_02_') === 0) return 'beliun1024';
  if (text.indexOf('롯백_03_') === 0) return 'beliun1023';
  if (text.indexOf('롯백_04_') === 0) return 'beliun1024';
  return '';
}

function rewriteDashboardMemo_v610_(memo, total) {
  var text = String(memo || '');
  var sent = '전송 ' + String(total || 0).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  if (/전송\s*[\d,]+/.test(text)) return text.replace(/전송\s*[\d,]+/, sent);
  return text;
}

function findHeaderIndex_v610_(headerRow, candidates) {
  var normalized = (headerRow || []).map(function(h) { return normalizeHeaderKey_v610_(h); });
  for (var i = 0; i < candidates.length; i++) {
    var key = normalizeHeaderKey_v610_(candidates[i]);
    var idx = normalized.indexOf(key);
    if (idx >= 0) return idx;
  }
  return -1;
}

function normalizeHeaderKey_v610_(v) {
  return String(v || '').replace(/\s+/g, '').replace(/\n/g, '').replace(/_/g, '').trim();
}
