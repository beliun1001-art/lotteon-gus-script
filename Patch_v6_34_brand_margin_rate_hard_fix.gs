/**
 * LOTTEON v6.34 brand margin rate hard fix
 *
 * 문제:
 * - 브랜드별_마진율 시트의 예상이익률(P열)이 0으로 표시되고 % 서식이 적용되지 않았습니다.
 * - v6.33의 전체 % 서식은 일반 헤더 기준으로 처리하지만, 브랜드별_마진율은 생성/서식 패치가 여러 번 겹치면서
 *   일부 율 컬럼이 다시 숫자 서식으로 덮이는 문제가 있었습니다.
 *
 * 수정:
 * - 브랜드별_마진율 시트만 별도로 강제 보정합니다.
 * - 마켓수수료율 = 마켓수수료/비용 ÷ 순수매출액
 * - 예상이익률 = 예상이익 ÷ 순수매출액
 * - 부가세반영이익률 = 부가세반영예상이익 ÷ 순수매출액
 * - 위 3개 율 컬럼은 0.0% 형식으로 고정합니다.
 * - 대시보드/부가세 생성/표시서식 정리/열너비 자동조정 후에도 마지막에 다시 적용합니다.
 */

var LOTTEON_PATCH_V634_BRAND_MARGIN_RATE_HARD_FIX_LOADED = true;

var __baseBuildBrandMarginSingleSource_v634 = typeof buildBrandMarginSingleSource_v628_ === 'function' ? buildBrandMarginSingleSource_v628_ : null;
var __baseApplyDisplayStandardsOnlyFast_v634 = typeof applyDisplayStandardsOnlyFast_v623 === 'function' ? applyDisplayStandardsOnlyFast_v623 : null;
var __baseApplyFastOutputSheetFormatting_v634 = typeof applyFastOutputSheetFormatting_v620_ === 'function' ? applyFastOutputSheetFormatting_v620_ : null;
var __baseRunColumnWidthStep_v634 = typeof runColumnWidthAutoAdjustStep_v623 === 'function' ? runColumnWidthAutoAdjustStep_v623 : null;
var __baseRefreshDashboardFastOnly_v634 = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;
var __baseGenerateVatReportsFullSeparated_v634 = typeof generateVatReportsFullSeparated_v622 === 'function' ? generateVatReportsFullSeparated_v622 : null;

buildBrandMarginSingleSource_v628_ = function(salesAgg) {
  var result = __baseBuildBrandMarginSingleSource_v634 ? __baseBuildBrandMarginSingleSource_v634.apply(this, arguments) : null;
  hardFixBrandMarginRateColumns_v634_();
  return result || { ok: true };
};

applyDisplayStandardsOnlyFast_v623 = function() {
  var result = __baseApplyDisplayStandardsOnlyFast_v634 ? __baseApplyDisplayStandardsOnlyFast_v634.apply(this, arguments) : null;
  hardFixBrandMarginRateColumns_v634_();
  try { SpreadsheetApp.getUi().alert('표시서식 정리 완료\n\n브랜드별_마진율의 마켓수수료율/예상이익률/부가세반영이익률을 0.0% 형식으로 재계산했습니다.'); } catch (e) {}
  return result || { ok: true };
};

applyFastOutputSheetFormatting_v620_ = function() {
  var result = __baseApplyFastOutputSheetFormatting_v634 ? __baseApplyFastOutputSheetFormatting_v634.apply(this, arguments) : null;
  hardFixBrandMarginRateColumns_v634_();
  return result || { ok: true };
};

runColumnWidthAutoAdjustStep_v623 = function() {
  var result = __baseRunColumnWidthStep_v634 ? __baseRunColumnWidthStep_v634.apply(this, arguments) : null;
  hardFixBrandMarginRateColumns_v634_();
  return result;
};

refreshDashboardFastOnly = function() {
  var result = __baseRefreshDashboardFastOnly_v634 ? __baseRefreshDashboardFastOnly_v634.apply(this, arguments) : null;
  hardFixBrandMarginRateColumns_v634_();
  return result;
};

generateVatReportsFullSeparated_v622 = function() {
  var result = __baseGenerateVatReportsFullSeparated_v634 ? __baseGenerateVatReportsFullSeparated_v634.apply(this, arguments) : null;
  hardFixBrandMarginRateColumns_v634_();
  return result;
};

function hardFixBrandMarginRateColumns_v634_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('브랜드별_마진율');
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return { skipped: true };

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return normalizeBrandMarginHeader_v634_(h); });

  var col = {
    netSales: findBrandMarginHeader_v634_(headers, ['순수매출액']),
    marketFee: findBrandMarginHeader_v634_(headers, ['마켓수수료/비용', '마켓수수료비용']),
    marketFeeRate: findBrandMarginHeader_v634_(headers, ['마켓수수료율']),
    profit: findBrandMarginHeader_v634_(headers, ['예상이익']),
    profitRate: findBrandMarginHeader_v634_(headers, ['예상이익률', '예상이익율']),
    vatProfit: findBrandMarginHeader_v634_(headers, ['부가세반영예상이익']),
    vatProfitRate: findBrandMarginHeader_v634_(headers, ['부가세반영이익률', '부가세반영이익율'])
  };

  var rowCount = lastRow - 1;
  var values = sheet.getRange(2, 1, rowCount, lastCol).getValues();
  var changed = false;

  for (var r = 0; r < values.length; r++) {
    var netSales = getCellNum_v634_(values[r], col.netSales);
    if (!netSales) continue;

    if (col.marketFeeRate >= 0 && col.marketFee >= 0) {
      var marketFeeRate = getCellNum_v634_(values[r], col.marketFee) / netSales;
      if (values[r][col.marketFeeRate] !== marketFeeRate) {
        values[r][col.marketFeeRate] = marketFeeRate;
        changed = true;
      }
    }

    if (col.profitRate >= 0 && col.profit >= 0) {
      var profitRate = getCellNum_v634_(values[r], col.profit) / netSales;
      if (values[r][col.profitRate] !== profitRate) {
        values[r][col.profitRate] = profitRate;
        changed = true;
      }
    }

    if (col.vatProfitRate >= 0 && col.vatProfit >= 0) {
      var vatProfitRate = getCellNum_v634_(values[r], col.vatProfit) / netSales;
      if (values[r][col.vatProfitRate] !== vatProfitRate) {
        values[r][col.vatProfitRate] = vatProfitRate;
        changed = true;
      }
    }
  }

  if (changed) sheet.getRange(2, 1, rowCount, lastCol).setValues(values);

  [col.marketFeeRate, col.profitRate, col.vatProfitRate].forEach(function(idx) {
    if (idx >= 0) {
      try {
        sheet.getRange(2, idx + 1, rowCount, 1)
          .setNumberFormat('0.0%')
          .setHorizontalAlignment('right');
      } catch (e) {}
    }
  });

  try {
    sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#d9eaf7').setHorizontalAlignment('center');
    log_('patch_v634_brand_margin_rate_fix', JSON.stringify(col));
  } catch (e) {}

  return { ok: true, columns: col };
}

function normalizeBrandMarginHeader_v634_(v) {
  return String(v == null ? '' : v).replace(/\n/g, '').replace(/\s+/g, '').trim();
}

function findBrandMarginHeader_v634_(headers, candidates) {
  for (var i = 0; i < headers.length; i++) {
    for (var j = 0; j < candidates.length; j++) {
      if (headers[i] === normalizeBrandMarginHeader_v634_(candidates[j])) return i;
    }
  }
  return -1;
}

function getCellNum_v634_(row, idx) {
  if (idx < 0) return 0;
  var v = row[idx];
  if (typeof toNumber_ === 'function') return toNumber_(v);
  if (typeof v === 'number') return v;
  var s = String(v == null ? '' : v).replace(/₩/g, '').replace(/,/g, '').replace(/%/g, '').trim();
  if (!s) return 0;
  var n = Number(s);
  return isNaN(n) ? 0 : n;
}
