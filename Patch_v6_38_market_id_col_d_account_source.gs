/**
 * LOTTEON v6.38 D column market id account source
 *
 * 확정 기준:
 * - 매출데이터_붙여넣기 시트의 D열 '마켓아이디'를 계정/사업자 구분의 단일 원천으로 사용합니다.
 * - 기존 원본계정ID/분석계정ID/계정번호/대표검색필터명 fallback보다 D열을 우선합니다.
 * - 사업자별_계정별_써머리, 부가세_신고자료, 미정산_쿠팡계정별 모두 D열 마켓아이디 기준으로 구분합니다.
 * - D열 값은 beliun 형식이 아니어도 그대로 계정 구분값으로 사용합니다.
 */

var LOTTEON_PATCH_V638_MARKET_ID_COL_D_ACCOUNT_SOURCE_LOADED = true;
var LOTTEON_V638_MARKET_ID_COLUMN_INDEX = 4; // D column, 1-based

var __basePrepareAccountSummarySalesAgg_v638 = typeof prepareAccountSummarySalesAgg_v637_ === 'function' ? prepareAccountSummarySalesAgg_v637_ : null;
var __baseMakeSourceRow_v638 = typeof makeSourceRow_v628_ === 'function' ? makeSourceRow_v628_ : null;
var __baseBuildAccountSummarySheet_v638 = typeof buildAccountSummarySheet_v637_ === 'function' ? buildAccountSummarySheet_v637_ : null;
var __baseBuildVatDetailSingleSource_v638 = typeof buildVatDetailSingleSource_v628_ === 'function' ? buildVatDetailSingleSource_v628_ : null;
var __baseGenerateVatReportsFullSeparated_v638 = typeof generateVatReportsFullSeparated_v622 === 'function' ? generateVatReportsFullSeparated_v622 : null;
var __baseRefreshDashboardFastOnly_v638 = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;
var __baseApplyDisplayStandardsOnlyFast_v638 = typeof applyDisplayStandardsOnlyFast_v623 === 'function' ? applyDisplayStandardsOnlyFast_v623 : null;

makeSourceRow_v628_ = function(sheetRow, obj, sales, purchase) {
  var row = __baseMakeSourceRow_v638 ? __baseMakeSourceRow_v638.apply(this, arguments) : {};
  var marketId = getMarketIdFromSourceRowNumber_v638_(sheetRow);
  if (marketId) row.accountId = marketId;
  return row;
};

// v6.36 미정산 계정 map도 D열 기준으로 재정의합니다.
sourceRowAccountMap_v636_ = function() {
  return buildMarketIdBySourceRowFromColumnD_v638_();
};

// v6.37 계정 써머리 계정 map도 D열 기준으로 재정의합니다.
buildAccountBySourceRowForSummary_v637_ = function() {
  return buildMarketIdBySourceRowFromColumnD_v638_();
};

prepareAccountSummarySalesAgg_v637_ = function(salesAgg) {
  salesAgg = __basePrepareAccountSummarySalesAgg_v638 ? __basePrepareAccountSummarySalesAgg_v638.apply(this, arguments) : (salesAgg || {});
  return forceSalesAggAccountFromColumnD_v638_(salesAgg);
};

buildAccountSummarySheet_v637_ = function(salesAgg) {
  salesAgg = forceSalesAggAccountFromColumnD_v638_(salesAgg || buildSingleSourceSalesAgg_v628_());
  return __baseBuildAccountSummarySheet_v638 ? __baseBuildAccountSummarySheet_v638.call(this, salesAgg) : null;
};

buildVatDetailSingleSource_v628_ = function(salesAgg) {
  salesAgg = forceSalesAggAccountFromColumnD_v638_(salesAgg || buildSingleSourceSalesAgg_v628_());
  return __baseBuildVatDetailSingleSource_v638 ? __baseBuildVatDetailSingleSource_v638.call(this, salesAgg) : null;
};

generateVatReportsFullSeparated_v622 = function() {
  var result = __baseGenerateVatReportsFullSeparated_v638 ? __baseGenerateVatReportsFullSeparated_v638.apply(this, arguments) : null;
  try {
    var salesAgg = forceSalesAggAccountFromColumnD_v638_(buildSingleSourceSalesAgg_v628_());
    buildAccountSummarySheet_v637_(salesAgg);
    buildVatDetailSingleSource_v628_(salesAgg);
    buildMarketIdAccountDiagnostic_v638_(salesAgg);
  } catch (e) { try { log_('patch_v638_after_vat_error', String(e && e.message ? e.message : e)); } catch(ignore) {} }
  return result || { ok: true };
};

refreshDashboardFastOnly = function() {
  var result = __baseRefreshDashboardFastOnly_v638 ? __baseRefreshDashboardFastOnly_v638.apply(this, arguments) : null;
  try {
    var salesAgg = forceSalesAggAccountFromColumnD_v638_(buildSingleSourceSalesAgg_v628_());
    buildAccountSummarySheet_v637_(salesAgg);
    buildMarketIdAccountDiagnostic_v638_(salesAgg);
  } catch (e) {}
  return result || { ok: true };
};

applyDisplayStandardsOnlyFast_v623 = function() {
  var result = __baseApplyDisplayStandardsOnlyFast_v638 ? __baseApplyDisplayStandardsOnlyFast_v638.apply(this, arguments) : null;
  try {
    var salesAgg = forceSalesAggAccountFromColumnD_v638_(buildSingleSourceSalesAgg_v628_());
    buildAccountSummarySheet_v637_(salesAgg);
    buildVatDetailSingleSource_v628_(salesAgg);
    buildMarketIdAccountDiagnostic_v638_(salesAgg);
    SpreadsheetApp.getUi().alert('표시서식 정리 완료\n\n계정/사업자 구분은 매출데이터_붙여넣기 D열 마켓아이디 기준으로 재적용했습니다.');
  } catch (e) {}
  return result || { ok: true };
};

function forceSalesAggAccountFromColumnD_v638_(salesAgg) {
  salesAgg = salesAgg || {};
  var map = buildMarketIdBySourceRowFromColumnD_v638_();
  var missing = 0;
  (salesAgg.detailRows || []).forEach(function(d) {
    var marketId = map[d.sourceRow] || '';
    if (marketId) d.accountId = marketId;
    else if (!d.accountId) { d.accountId = '마켓아이디없음'; missing++; }
  });
  salesAgg.__v638MarketIdColumnDApplied = true;
  salesAgg.__v638MarketIdMissingCount = missing;
  return salesAgg;
}

function buildMarketIdBySourceRowFromColumnD_v638_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SALES_IN) || ss.getSheetByName('매출데이터_붙여넣기');
  var out = {};
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < LOTTEON_V638_MARKET_ID_COLUMN_INDEX) return out;
  var rowCount = sheet.getLastRow() - 1;
  var values = sheet.getRange(2, LOTTEON_V638_MARKET_ID_COLUMN_INDEX, rowCount, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    var marketId = normalizeMarketId_v638_(values[i][0]);
    if (marketId) out[i + 2] = marketId;
  }
  return out;
}

function getMarketIdFromSourceRowNumber_v638_(sourceRowNo) {
  var map = buildMarketIdBySourceRowFromColumnD_v638_();
  return map[sourceRowNo] || '';
}

function buildMarketIdAccountDiagnostic_v638_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('계정구분_D열검증') || ss.insertSheet('계정구분_D열검증');
  var map = buildMarketIdBySourceRowFromColumnD_v638_();
  var counts = {};
  var missing = 0;
  (salesAgg.detailRows || []).forEach(function(d) {
    var id = map[d.sourceRow] || d.accountId || '마켓아이디없음';
    counts[id] = (counts[id] || 0) + 1;
    if (!map[d.sourceRow]) missing++;
  });
  var rows = Object.keys(counts).sort().map(function(k) { return [k, counts[k], 'D열 마켓아이디 기준']; });
  rows.unshift(['분석대상행수', (salesAgg.detailRows || []).length, '']);
  rows.unshift(['D열 마켓아이디 없음', missing, 'sourceRow 기준 D열 공란']);
  rows.unshift(['원천기준', '매출데이터_붙여넣기 D열 마켓아이디', 'v6.38']);
  try {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 3).setValues([['항목/마켓아이디','행수','메모']]);
    if (rows.length) sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    sheet.getRange(1, 1, 1, 3).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
    sheet.autoResizeColumns(1, 3);
  } catch (e) {}
}

function normalizeMarketId_v638_(v) {
  return String(v == null ? '' : v).replace(/,/g, '').trim();
}
