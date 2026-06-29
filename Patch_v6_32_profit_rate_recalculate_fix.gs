/**
 * LOTTEON v6.32 profit rate recalculate fix
 *
 * 문제:
 * - 대시보드 요약의 예상이익은 정상인데 예상이익률이 0.00%로 표시되었습니다.
 * - 원인은 salesAgg.totalEstimatedProfitRate 값이 일부 경로에서 0 또는 미갱신 상태로 남았기 때문입니다.
 *
 * 수정:
 * - 대시보드 생성 직전 전체 이익률을 순수매출액/예상이익 기준으로 강제 재계산합니다.
 * - 브랜드별 이익률도 브랜드 순수매출액/브랜드 예상이익 기준으로 재계산합니다.
 * - 대시보드 시트에 이미 생성된 요약 예상이익률 셀도 직접 보정합니다.
 */

var LOTTEON_PATCH_V632_PROFIT_RATE_RECALCULATE_FIX_LOADED = true;

var __baseRebuildDashboardNoFilterAccount_v632 = typeof rebuildDashboardNoFilterAccount_v630_ === 'function' ? rebuildDashboardNoFilterAccount_v630_ : null;
var __baseBuildBrandMarginSingleSource_v632 = typeof buildBrandMarginSingleSource_v628_ === 'function' ? buildBrandMarginSingleSource_v628_ : null;
var __baseBuildFinancialValidationSheets_v632 = typeof buildFinancialValidationSheets_v628_ === 'function' ? buildFinancialValidationSheets_v628_ : null;
var __baseHardFixDashboardDisplay_v632 = typeof hardFixDashboardDisplay_v630_ === 'function' ? hardFixDashboardDisplay_v630_ : null;

rebuildDashboardNoFilterAccount_v630_ = function(salesAgg, filterAgg) {
  salesAgg = recalculateProfitRates_v632_(salesAgg || (typeof buildSingleSourceSalesAgg_v628_ === 'function' ? buildSingleSourceSalesAgg_v628_() : null));
  var result = __baseRebuildDashboardNoFilterAccount_v632 ? __baseRebuildDashboardNoFilterAccount_v632.call(this, salesAgg, filterAgg) : null;
  forceDashboardProfitRateCell_v632_();
  return result;
};

rebuildDashboardSingleSource_v628_ = function(salesAgg, filterAgg) {
  salesAgg = recalculateProfitRates_v632_(salesAgg || (typeof buildSingleSourceSalesAgg_v628_ === 'function' ? buildSingleSourceSalesAgg_v628_() : null));
  return rebuildDashboardNoFilterAccount_v630_(salesAgg, filterAgg);
};

buildBrandMarginSingleSource_v628_ = function(salesAgg) {
  salesAgg = recalculateProfitRates_v632_(salesAgg || (typeof buildSingleSourceSalesAgg_v628_ === 'function' ? buildSingleSourceSalesAgg_v628_() : null));
  return __baseBuildBrandMarginSingleSource_v632 ? __baseBuildBrandMarginSingleSource_v632.call(this, salesAgg) : null;
};

buildFinancialValidationSheets_v628_ = function(salesAgg) {
  salesAgg = recalculateProfitRates_v632_(salesAgg || (typeof buildSingleSourceSalesAgg_v628_ === 'function' ? buildSingleSourceSalesAgg_v628_() : null));
  return __baseBuildFinancialValidationSheets_v632 ? __baseBuildFinancialValidationSheets_v632.call(this, salesAgg) : null;
};

hardFixDashboardDisplay_v630_ = function() {
  var result = __baseHardFixDashboardDisplay_v632 ? __baseHardFixDashboardDisplay_v632.apply(this, arguments) : null;
  forceDashboardProfitRateCell_v632_();
  return result || { ok: true };
};

function recalculateProfitRates_v632_(salesAgg) {
  if (!salesAgg) return salesAgg;

  var netSales = toNum_v632_(salesAgg.totalNetSales);
  var profit = toNum_v632_(salesAgg.totalEstimatedProfit);
  if (!profit && salesAgg.totalSettlementBasis != null && salesAgg.totalPurchase != null) {
    profit = toNum_v632_(salesAgg.totalSettlementBasis) - toNum_v632_(salesAgg.totalPurchase);
    salesAgg.totalEstimatedProfit = profit;
    salesAgg.totalNetProfit = profit;
  }
  salesAgg.totalEstimatedProfitRate = netSales ? profit / netSales : 0;
  salesAgg.totalNetProfitRate = salesAgg.totalEstimatedProfitRate;
  salesAgg.totalMarginRate = salesAgg.totalEstimatedProfitRate;

  Object.keys(salesAgg.byBrand || {}).forEach(function(key) {
    var b = salesAgg.byBrand[key];
    var bNet = toNum_v632_(b.netSalesAmount || b.salesAmount);
    var bProfit = toNum_v632_(b.estimatedProfitAmount || b.netProfitAmount || b.marginAmount);
    if (!bProfit && b.settlementBasisAmount != null && b.purchaseAmount != null) {
      bProfit = toNum_v632_(b.settlementBasisAmount) - toNum_v632_(b.purchaseAmount);
      b.estimatedProfitAmount = bProfit;
      b.netProfitAmount = bProfit;
      b.marginAmount = bProfit;
    }
    b.estimatedProfitRate = bNet ? bProfit / bNet : 0;
    b.netProfitRate = b.estimatedProfitRate;
    b.marginRate = b.estimatedProfitRate;
  });

  return salesAgg;
}

function forceDashboardProfitRateCell_v632_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!sheet || sheet.getLastRow() < 2) return { skipped: true };

  var lastRow = Math.min(sheet.getLastRow(), 40);
  var values = sheet.getRange(1, 1, lastRow, Math.min(sheet.getLastColumn(), 16)).getValues();
  var netSales = 0;
  var profit = 0;
  var rateRow = 0;

  for (var r = 0; r < values.length; r++) {
    var item = String(values[r][1] == null ? '' : values[r][1]).replace(/\n/g, '').replace(/\s+/g, '').trim();
    if (item === '순수매출액') netSales = toNum_v632_(values[r][2]);
    if (item === '예상이익') profit = toNum_v632_(values[r][2]);
    if (item === '예상이익률') rateRow = r + 1;
  }

  if (rateRow && netSales) {
    try {
      sheet.getRange(rateRow, 3)
        .setValue(profit / netSales)
        .setNumberFormat('0.00%')
        .setHorizontalAlignment('right');
      try { log_('patch_v632_profit_rate_fix', 'profit=' + profit + ' netSales=' + netSales + ' rate=' + (profit / netSales)); } catch (e) {}
    } catch (e) {}
  }
  return { ok: true };
}

function toNum_v632_(v) {
  if (typeof toNumber_ === 'function') return toNumber_(v);
  if (typeof v === 'number') return v;
  var s = String(v == null ? '' : v).replace(/₩/g, '').replace(/,/g, '').replace(/%/g, '').trim();
  if (!s) return 0;
  var n = Number(s);
  return isNaN(n) ? 0 : n;
}
