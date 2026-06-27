/**
 * LOTTEON v6.09 dashboard brand/filter sync patch
 *
 * 목적:
 * - 대시보드의 확대TOP / 상품갈이 / 조치필요 영역에서 브랜드명과 대표검색필터명을
 *   필터별_상품수 기준 최신 값으로 보정합니다.
 * - 브랜드명은 v6.08 기준과 동일하게 관리용 앞 숫자_ 및 뒤 _번호를 제거합니다.
 * - 대표검색필터명은 실제 필터별_상품수.검색필터명 exact 값을 사용합니다.
 *
 * 예:
 * - K2 / 롯백_04_K2 -> K2 / 롯백_04_1_K2
 * - 숲 / 롯백_03_숲 -> 숲 / 롯백_01_5_숲 또는 전송수와 일치하는 실제 필터
 * - 1K2 / 롯백_04_1K2 -> K2 / 롯백_04_1_K2
 * - 플라스틱아일랜드 / 롯백_03_플라스틱아일랜드 -> 플라스틱아일랜드 / 롯백_03_3_플라스틱아일랜드_01
 */

var LOTTEON_PATCH_V609_DASHBOARD_BRAND_FILTER_SYNC_LOADED = true;

var __baseBuildDashboard_v609 = typeof buildDashboard_ === 'function' ? buildDashboard_ : null;
var __baseRefreshDashboardFastOnly_v609 = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;
var __baseRunPendingChangesApproval_v609 = typeof runPendingChangesApproval === 'function' ? runPendingChangesApproval : null;
var __baseRefreshCoreSummaryAndDashboard_v609 = typeof refreshCoreSummaryAndDashboardWithRetransmitLogDates === 'function' ? refreshCoreSummaryAndDashboardWithRetransmitLogDates : null;

if (__baseBuildDashboard_v609) {
  buildDashboard_ = function() {
    var result = __baseBuildDashboard_v609.apply(this, arguments);
    patchDashboardBrandAndFilterFromFilterSheet_v609_();
    return result;
  };
}

refreshDashboardFastOnly = function() {
  var result = __baseRefreshDashboardFastOnly_v609 ? __baseRefreshDashboardFastOnly_v609.apply(this, arguments) : null;
  patchDashboardBrandAndFilterFromFilterSheet_v609_();
  return result;
};

runPendingChangesApproval = function() {
  var result = __baseRunPendingChangesApproval_v609 ? __baseRunPendingChangesApproval_v609.apply(this, arguments) : null;
  patchDashboardBrandAndFilterFromFilterSheet_v609_();
  return result;
};

refreshCoreSummaryAndDashboardWithRetransmitLogDates = function() {
  var result = __baseRefreshCoreSummaryAndDashboard_v609 ? __baseRefreshCoreSummaryAndDashboard_v609.apply(this, arguments) : null;
  patchDashboardBrandAndFilterFromFilterSheet_v609_();
  return result;
};

function patchDashboardBrandAndFilterFromFilterSheet_v609_() {
  var ss = SpreadsheetApp.getActive();
  var dashboard = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!dashboard || dashboard.getLastRow() < 2) return { updated: 0, reason: 'NO_DASHBOARD' };

  var map = buildDashboardFilterLookup_v609_();
  if (!map || !map.rows || !map.rows.length) return { updated: 0, reason: 'NO_FILTER_MAP' };

  var range = dashboard.getDataRange();
  var values = range.getValues();
  if (!values || values.length < 2) return { updated: 0, reason: 'EMPTY_DASHBOARD' };

  var updated = 0;
  var targetGroups = { '확대TOP': true, '상품갈이': true, '조치필요': true };

  for (var r = 0; r < values.length; r++) {
    var group = String(values[r][0] || '').trim();
    if (!targetGroups[group]) continue;

    var currentBrand = String(values[r][1] || '').trim();
    var currentFilter = String(values[r][2] || '').trim();
    var currentAccountId = String(values[r][3] || '').trim();
    var currentTotal = toNumber_(values[r][4]);
    var salesProduct = toNumber_(values[r][5]);

    var canonical = canonicalDashboardBrand_v609_(currentBrand) || canonicalBrandFromFilterName_v609_(currentFilter);
    var item = chooseDashboardFilterItem_v609_(map, canonical, currentFilter, currentAccountId, currentTotal);
    if (!item) continue;

    var rowChanged = false;
    if (values[r][1] !== item.brand) { values[r][1] = item.brand; rowChanged = true; }
    if (values[r][2] !== item.filterName) { values[r][2] = item.filterName; rowChanged = true; }
    if (values[r][3] !== item.accountId) { values[r][3] = item.accountId; rowChanged = true; }
    if (values[r][4] !== item.total) { values[r][4] = item.total; rowChanged = true; }

    if (item.total > 0 && values[r].length > 6) {
      var newRate = salesProduct / item.total;
      if (values[r][6] !== newRate) { values[r][6] = newRate; rowChanged = true; }
    }

    if (values[r].length > 10) {
      var memo = String(values[r][10] || '');
      var newMemo = patchDashboardMemoSentCount_v609_(memo, item.total);
      if (newMemo !== memo) { values[r][10] = newMemo; rowChanged = true; }
    }

    if (rowChanged) updated++;
  }

  if (updated > 0) {
    range.setValues(values);
    try {
      dashboard.getRange(1, 5, dashboard.getLastRow(), 1).setNumberFormat('#,##0');
      dashboard.getRange(1, 7, dashboard.getLastRow(), 1).setNumberFormat('0.00%');
    } catch (e) {}
  }

  try {
    log_('patch_dashboard_brand_filter_v609', 'updated=' + updated + ' / source=필터별_상품수');
  } catch (e) {}

  return { updated: updated };
}

function buildDashboardFilterLookup_v609_() {
  var result = { rows: [], byFilter: {}, byBrand: {} };
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.FILTERS);
  if (!sheet || sheet.getLastRow() < 2) return result;

  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var colFilter = findHeaderIndex_v609_(header, ['검색필터명']);
  var colBrand = findHeaderIndex_v609_(header, ['브랜드명']);
  var colTotal = findHeaderIndex_v609_(header, ['API_totalCount', 'APItotalCount']);
  var colAccountId = findHeaderIndex_v609_(header, ['쿠팡계정ID']);
  if (colFilter < 0 || colTotal < 0) return result;

  for (var r = 1; r < values.length; r++) {
    var filterName = String(values[r][colFilter] || '').trim();
    if (!filterName || filterName.indexOf(CONFIG.FILTER_PREFIX) !== 0) continue;

    var rawBrand = colBrand >= 0 ? String(values[r][colBrand] || '').trim() : '';
    var brand = canonicalDashboardBrand_v609_(rawBrand) || canonicalBrandFromFilterName_v609_(filterName);
    var total = toNumber_(values[r][colTotal]);
    var accountId = colAccountId >= 0 ? String(values[r][colAccountId] || '').trim() : accountIdFromFilterName_v609_(filterName);
    var item = {
      filterName: filterName,
      brand: brand,
      key: normalizeDashboardBrandKey_v609_(brand),
      total: total,
      accountId: accountId
    };

    result.rows.push(item);
    result.byFilter[filterName] = item;
    if (item.key) {
      if (!result.byBrand[item.key]) result.byBrand[item.key] = [];
      result.byBrand[item.key].push(item);
    }
  }

  return result;
}

function chooseDashboardFilterItem_v609_(map, canonicalBrand, currentFilter, currentAccountId, currentTotal) {
  if (currentFilter && map.byFilter[currentFilter]) return map.byFilter[currentFilter];

  var key = normalizeDashboardBrandKey_v609_(canonicalBrand);
  var list = key ? (map.byBrand[key] || []) : [];
  if (!list.length) return null;

  var exactTotal = list.filter(function(item) { return item.total === currentTotal; });
  if (exactTotal.length === 1) return exactTotal[0];
  if (exactTotal.length > 1 && currentAccountId) {
    var exactTotalAccount = exactTotal.filter(function(item) { return item.accountId === currentAccountId; });
    if (exactTotalAccount.length) return exactTotalAccount[0];
  }
  if (exactTotal.length > 1) return exactTotal[0];

  if (currentAccountId) {
    var accountMatches = list.filter(function(item) { return item.accountId === currentAccountId; });
    if (accountMatches.length === 1) return accountMatches[0];
    if (accountMatches.length > 1) return chooseLargestTotalItem_v609_(accountMatches);
  }

  return chooseLargestTotalItem_v609_(list);
}

function chooseLargestTotalItem_v609_(list) {
  return (list || []).slice().sort(function(a, b) {
    if ((b.total || 0) !== (a.total || 0)) return (b.total || 0) - (a.total || 0);
    return String(a.filterName || '').localeCompare(String(b.filterName || ''));
  })[0] || null;
}

function canonicalBrandFromFilterName_v609_(filterName) {
  var text = String(filterName || '').trim();
  if (!text) return '';
  var first = text.indexOf('_');
  if (first < 0) return '';
  var second = text.indexOf('_', first + 1);
  if (second < 0) return '';
  return canonicalDashboardBrand_v609_(text.slice(second + 1).trim());
}

function canonicalDashboardBrand_v609_(brandName) {
  return String(brandName || '')
    .trim()
    .replace(/^\d+_?/, '')
    .replace(/_\d+$/, '')
    .trim();
}

function normalizeDashboardBrandKey_v609_(brandName) {
  return canonicalDashboardBrand_v609_(brandName).toLowerCase().replace(/\s+/g, '').replace(/[\[\]\(\)\{\}\-_]/g, '').trim();
}

function accountIdFromFilterName_v609_(filterName) {
  var text = String(filterName || '').trim();
  if (text.indexOf('롯백_01_') === 0) return 'beliun1021';
  if (text.indexOf('롯백_02_') === 0) return 'beliun1024';
  if (text.indexOf('롯백_03_') === 0) return 'beliun1023';
  if (text.indexOf('롯백_04_') === 0) return 'beliun1024';
  return '';
}

function patchDashboardMemoSentCount_v609_(memo, total) {
  var text = String(memo || '');
  if (!text) return text;
  var sent = '전송 ' + String(total || 0).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  if (/전송\s*[\d,]+/.test(text)) return text.replace(/전송\s*[\d,]+/, sent);
  return text;
}

function findHeaderIndex_v609_(headerRow, candidates) {
  var normalized = (headerRow || []).map(function(h) { return normalizeHeaderKey_v609_(h); });
  for (var i = 0; i < candidates.length; i++) {
    var key = normalizeHeaderKey_v609_(candidates[i]);
    var idx = normalized.indexOf(key);
    if (idx >= 0) return idx;
  }
  return -1;
}

function normalizeHeaderKey_v609_(v) {
  return String(v || '').replace(/\s+/g, '').replace(/\n/g, '').replace(/_/g, '').trim();
}
