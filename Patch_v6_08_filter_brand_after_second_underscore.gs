/**
 * LOTTEON v6.08 filter brand name patch
 *
 * 목적:
 * - 필터별_상품수 시트의 브랜드명은 검색필터명에서 두 번째 '_' 다음 전체 문자열을 사용합니다.
 * - 예: 롯백_04_1_K2 -> 1_K2
 * - 예: 롯백_03_3_플라스틱아일랜드_01 -> 3_플라스틱아일랜드_01
 *
 * 주의:
 * - 이 패치는 필터별_상품수 시트의 브랜드명 표시 기준만 보정합니다.
 * - API_totalCount 원본 값은 수정하지 않습니다.
 */

var LOTTEON_PATCH_V608_FILTER_BRAND_AFTER_SECOND_UNDERSCORE_LOADED = true;

var __baseRunFilterListResumeBatch_v608 = typeof runFilterListResumeBatch_ === 'function' ? runFilterListResumeBatch_ : null;
var __baseRefreshFilterCountsFast_v608 = typeof refreshFilterCountsFast === 'function' ? refreshFilterCountsFast : null;
var __baseRunDailyFilterCountsStep_v608 = typeof runDailyFilterCountsStep_ === 'function' ? runDailyFilterCountsStep_ : null;
var __baseRunDailyFilterCountsOnceManual_v608 = typeof runDailyFilterCountsOnceManual === 'function' ? runDailyFilterCountsOnceManual : null;
var __baseRefreshDashboardFastOnly_v608 = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;

if (__baseRunFilterListResumeBatch_v608) {
  runFilterListResumeBatch_ = function() {
    var result = __baseRunFilterListResumeBatch_v608.apply(this, arguments);
    normalizeFilterSheetBrandNamesFromSecondUnderscore_v608_();
    return result;
  };
}

if (__baseRefreshFilterCountsFast_v608) {
  refreshFilterCountsFast = function() {
    var result = __baseRefreshFilterCountsFast_v608.apply(this, arguments);
    normalizeFilterSheetBrandNamesFromSecondUnderscore_v608_();
    return result;
  };
}

if (__baseRunDailyFilterCountsStep_v608) {
  runDailyFilterCountsStep_ = function() {
    var result = __baseRunDailyFilterCountsStep_v608.apply(this, arguments);
    normalizeFilterSheetBrandNamesFromSecondUnderscore_v608_();
    return result;
  };
}

if (__baseRunDailyFilterCountsOnceManual_v608) {
  runDailyFilterCountsOnceManual = function() {
    var result = __baseRunDailyFilterCountsOnceManual_v608.apply(this, arguments);
    normalizeFilterSheetBrandNamesFromSecondUnderscore_v608_();
    return result;
  };
}

if (__baseRefreshDashboardFastOnly_v608) {
  refreshDashboardFastOnly = function() {
    normalizeFilterSheetBrandNamesFromSecondUnderscore_v608_();
    return __baseRefreshDashboardFastOnly_v608.apply(this, arguments);
  };
}

function normalizeFilterSheetBrandNamesFromSecondUnderscore_v608_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.FILTERS);
  if (!sheet || sheet.getLastRow() < 2) return { updated: 0, reason: 'NO_FILTER_SHEET' };

  var values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return { updated: 0, reason: 'EMPTY_FILTER_SHEET' };

  var header = values[0];
  var colFilter = findHeaderIndex_v608_(header, ['검색필터명']);
  var colBrand = findHeaderIndex_v608_(header, ['브랜드명']);
  if (colFilter < 0 || colBrand < 0) return { updated: 0, reason: 'HEADER_NOT_FOUND' };

  var updated = 0;
  var brandValues = [];
  for (var r = 1; r < values.length; r++) {
    var filterName = String(values[r][colFilter] || '').trim();
    var newBrand = brandNameAfterSecondUnderscore_v608_(filterName);
    if (!newBrand) newBrand = String(values[r][colBrand] || '').trim();
    brandValues.push([newBrand]);
    if (String(values[r][colBrand] || '').trim() !== newBrand) updated++;
  }

  if (brandValues.length) {
    sheet.getRange(2, colBrand + 1, brandValues.length, 1).setValues(brandValues);
  }

  try {
    log_('patch_filter_brand_second_underscore_v608', 'updated=' + updated + ' / rule=검색필터명 두 번째 underscore 다음');
  } catch (e) {}

  return { updated: updated };
}

function brandNameAfterSecondUnderscore_v608_(filterName) {
  var text = String(filterName || '').trim();
  if (!text) return '';
  var first = text.indexOf('_');
  if (first < 0) return '';
  var second = text.indexOf('_', first + 1);
  if (second < 0) return '';
  return text.slice(second + 1).trim();
}

function findHeaderIndex_v608_(headerRow, candidates) {
  var normalized = (headerRow || []).map(function(h) { return normalizeHeaderKey_v608_(h); });
  for (var i = 0; i < candidates.length; i++) {
    var key = normalizeHeaderKey_v608_(candidates[i]);
    var idx = normalized.indexOf(key);
    if (idx >= 0) return idx;
  }
  return -1;
}

function normalizeHeaderKey_v608_(v) {
  return String(v || '').replace(/\s+/g, '').replace(/\n/g, '').replace(/_/g, '').trim();
}
