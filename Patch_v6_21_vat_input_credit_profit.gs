/**
 * LOTTEON v6.21 VAT input credit / profit reference patch
 *
 * 확정 기준:
 * - 부가세 신고 관점: 납부예상부가세 = 매출부가세 - 공제가능 매입부가세
 * - 매출부가세 = 순수매출액 - ROUND(순수매출액 / 1.1)
 * - 매입부가세 = 매입금액 - ROUND(매입금액 / 1.1)
 * - 매입금액이 부가세 포함 금액이라는 전제입니다.
 * - 공제 불가 매입이 섞여 있으면 실제 신고 전 별도 조정이 필요합니다.
 *
 * 출력:
 * - 브랜드별_마진율에 매출부가세/매입부가세/납부예상부가세/부가세반영예상이익 추가
 * - 부가세_신고자료, 부가세_상품별, 부가세_고객별, 부가세_주문번호별에 매입부가세와 납부예상부가세 추가
 */

var LOTTEON_PATCH_V621_VAT_INPUT_CREDIT_PROFIT_LOADED = true;

buildBrandMarginSheet_v611_ = function(salesAgg) {
  return buildBrandMarginSheetVatCredit_v621_(salesAgg || aggregateSalesByBrand_v611_());
};

buildVatReportSheet_v611_ = function(salesAgg) {
  return buildVatBreakdownSheetsVatCredit_v621_(salesAgg || aggregateSalesByBrand_v611_());
};

function buildBrandMarginSheetVatCredit_v621_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('브랜드별_마진율') || ss.insertSheet('브랜드별_마진율');
  var headers = [
    '브랜드명','주문건수','고객수','판매수량','매출상품수',
    '총매출액','취소/반품매출','순수매출액','실제정산금액','예상정산금액(미정산)','정산기준금액',
    '마켓수수료/비용','마켓수수료율','매입금액','예상이익','예상이익률',
    '매출부가세','매입부가세','납부예상부가세','부가세반영예상이익','부가세반영이익률',
    '미정산주문건수','30일초과미정산건수','비고'
  ];

  var rows = Object.keys(salesAgg.byBrand || {}).map(function(key) {
    var b = salesAgg.byBrand[key];
    var netSales = b.netSalesAmount || b.salesAmount || 0;
    var purchase = b.purchaseAmount || 0;
    var salesVat = vatPartFromVatIncluded_v621_(netSales);
    var purchaseVat = vatPartFromVatIncluded_v621_(purchase);
    var payableVat = salesVat - purchaseVat;
    var estimatedProfit = b.estimatedProfitAmount || b.netProfitAmount || b.marginAmount || 0;
    var afterVatProfit = estimatedProfit - payableVat;
    var afterVatRate = netSales ? afterVatProfit / netSales : 0;
    return [
      b.brand,
      b.orderCount || 0,
      b.customerCount || 0,
      b.quantity || 0,
      b.salesProductCount || 0,
      b.grossSalesAmount || 0,
      b.cancelSalesAmount || 0,
      netSales,
      b.actualSettlementAmount || 0,
      b.estimatedSettlementAmount || 0,
      b.settlementBasisAmount || b.settlementAmount || 0,
      b.marketFeeAmount || 0,
      b.marketFeeRate || 0,
      purchase,
      estimatedProfit,
      b.estimatedProfitRate || b.netProfitRate || b.marginRate || 0,
      salesVat,
      purchaseVat,
      payableVat,
      afterVatProfit,
      afterVatRate,
      b.unsettledOrderCount || 0,
      b.overdueUnsettledOrderCount || 0,
      '납부예상부가세=매출부가세-매입부가세 / 부가세반영예상이익=예상이익-납부예상부가세'
    ];
  }).sort(function(a, b) { return b[7] - a[7]; });

  replaceSheetValues_v611_(sheet, headers, rows);
  formatNumberColumns_v611_(sheet, [2,3,4,5,22,23]);
  formatMoneyColumns_v611_(sheet, [6,7,8,9,10,11,12,14,15,17,18,19,20]);
  if (rows.length) {
    sheet.getRange(2, 13, rows.length, 1).setNumberFormat('0.00%');
    sheet.getRange(2, 16, rows.length, 1).setNumberFormat('0.00%');
    sheet.getRange(2, 21, rows.length, 1).setNumberFormat('0.00%');
  }
  if (typeof applyFastOutputSheetFormatting_v620_ === 'function') applyFastOutputSheetFormatting_v620_();
  return { rows: rows.length };
}

function buildVatBreakdownSheetsVatCredit_v621_(salesAgg) {
  buildVatDetailSheetVatCredit_v621_(salesAgg);
  buildVatProductSheetVatCredit_v621_(salesAgg);
  buildVatCustomerSheetVatCredit_v621_(salesAgg);
  buildVatOrderSheetVatCredit_v621_(salesAgg);
  if (typeof buildUnsettledSettlementByAccountSheet_v616_ === 'function') buildUnsettledSettlementByAccountSheet_v616_(salesAgg);
  if (typeof applyFastOutputSheetFormatting_v620_ === 'function') applyFastOutputSheetFormatting_v620_();
  return null;
}

function buildVatDetailSheetVatCredit_v621_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('부가세_신고자료') || ss.insertSheet('부가세_신고자료');
  var headers = [
    '년','월','일','주문번호','고객명','브랜드명','상품번호','상품명','판매수량',
    '순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료/비용',
    '매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','비고'
  ];
  var rows = (salesAgg.detailRows || []).map(function(d) {
    var salesSplit = splitVatIncluded_v621_(d.netSales || 0);
    var purchaseSplit = splitVatIncluded_v621_(d.purchase || 0);
    var payableVat = salesSplit.vat - purchaseSplit.vat;
    var profit = d.estimatedProfit != null ? d.estimatedProfit : (d.netProfit || 0);
    return [
      d.year, d.month, d.day, d.orderNo, d.customer, d.brand, d.productNo, d.productName, d.qty,
      d.netSales || 0, salesSplit.supply, salesSplit.vat, d.settlementBasis || d.settlement || 0, d.marketFee || 0,
      d.purchase || 0, purchaseSplit.supply, purchaseSplit.vat, payableVat, profit, profit - payableVat,
      (d.note || d.unsettledStatus || '') + (d.overdueStatus ? ' / ' + d.overdueStatus : '')
    ];
  }).sort(sortVatRows_v613_);
  replaceSheetValues_v611_(sheet, headers, rows);
  formatNumberColumns_v611_(sheet, [1,2,3,9]);
  formatMoneyColumns_v611_(sheet, [10,11,12,13,14,15,16,17,18,19,20]);
}

function buildVatProductSheetVatCredit_v621_(salesAgg) {
  var grouped = groupVatDetailsForVatCredit_v621_(salesAgg.detailRows || [], function(d) {
    return [d.brand, d.productNo, d.productName].join('|');
  }, function(d) {
    return { brand: d.brand, productNo: d.productNo, productName: d.productName, orderSet: {}, customerSet: {} };
  });
  Object.keys(grouped).forEach(function(k) {
    var g = grouped[k];
    (g.rows || []).forEach(function(d) { if (d.orderNo) g.orderSet[d.orderNo] = true; if (d.customer) g.customerSet[d.customer] = true; });
  });
  var rows = Object.keys(grouped).map(function(k) {
    var g = grouped[k];
    var salesSplit = splitVatIncluded_v621_(g.netSales);
    var purchaseSplit = splitVatIncluded_v621_(g.purchase);
    var payableVat = salesSplit.vat - purchaseSplit.vat;
    return [g.brand, g.productNo, g.productName, g.qty, g.netSales, salesSplit.supply, salesSplit.vat, g.purchase, purchaseSplit.supply, purchaseSplit.vat, payableVat, g.profit, g.profit - payableVat, Object.keys(g.orderSet).length, Object.keys(g.customerSet).length];
  }).sort(function(a, b) { return b[4] - a[4]; });
  writeVatSummarySheet_v613_('부가세_상품별', ['브랜드명','상품번호','상품명','판매수량','순수매출액','매출공급가액','매출부가세','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','주문건수','고객수'], rows, [5,6,7,8,9,10,11,12,13], [4,14,15]);
}

function buildVatCustomerSheetVatCredit_v621_(salesAgg) {
  var grouped = groupVatDetailsForVatCredit_v621_(salesAgg.detailRows || [], function(d) {
    return d.customer || '고객명미확인';
  }, function(d) {
    return { customer: d.customer || '고객명미확인', orderSet: {}, productSet: {} };
  });
  Object.keys(grouped).forEach(function(k) {
    var g = grouped[k];
    (g.rows || []).forEach(function(d) { if (d.orderNo) g.orderSet[d.orderNo] = true; if (d.productNo || d.productName) g.productSet[(d.productNo || '') + '|' + (d.productName || '')] = true; });
  });
  var rows = Object.keys(grouped).map(function(k) {
    var g = grouped[k];
    var salesSplit = splitVatIncluded_v621_(g.netSales);
    var purchaseSplit = splitVatIncluded_v621_(g.purchase);
    var payableVat = salesSplit.vat - purchaseSplit.vat;
    return [g.customer, Object.keys(g.orderSet).length, Object.keys(g.productSet).length, g.qty, g.netSales, salesSplit.supply, salesSplit.vat, g.purchase, purchaseSplit.supply, purchaseSplit.vat, payableVat, g.profit, g.profit - payableVat];
  }).sort(function(a, b) { return b[4] - a[4]; });
  writeVatSummarySheet_v613_('부가세_고객별', ['고객명','주문건수','상품수','판매수량','순수매출액','매출공급가액','매출부가세','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익'], rows, [5,6,7,8,9,10,11,12,13], [2,3,4]);
}

function buildVatOrderSheetVatCredit_v621_(salesAgg) {
  var grouped = groupVatDetailsForVatCredit_v621_(salesAgg.detailRows || [], function(d) {
    return d.orderNo || '주문번호미확인';
  }, function(d) {
    return { year: d.year, month: d.month, day: d.day, orderNo: d.orderNo || '주문번호미확인', customer: d.customer, brandSet: {}, productSet: {} };
  });
  Object.keys(grouped).forEach(function(k) {
    var g = grouped[k];
    (g.rows || []).forEach(function(d) { if (d.brand) g.brandSet[d.brand] = true; if (d.productNo || d.productName) g.productSet[(d.productNo || '') + '|' + (d.productName || '')] = true; if (!g.customer && d.customer) g.customer = d.customer; });
  });
  var rows = Object.keys(grouped).map(function(k) {
    var g = grouped[k];
    var salesSplit = splitVatIncluded_v621_(g.netSales);
    var purchaseSplit = splitVatIncluded_v621_(g.purchase);
    var payableVat = salesSplit.vat - purchaseSplit.vat;
    return [g.year, g.month, g.day, g.orderNo, g.customer, Object.keys(g.brandSet).length, Object.keys(g.productSet).length, g.qty, g.netSales, salesSplit.supply, salesSplit.vat, g.purchase, purchaseSplit.supply, purchaseSplit.vat, payableVat, g.profit, g.profit - payableVat];
  }).sort(sortVatRows_v613_);
  writeVatSummarySheet_v613_('부가세_주문번호별', ['년','월','일','주문번호','고객명','브랜드수','상품수','판매수량','순수매출액','매출공급가액','매출부가세','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익'], rows, [9,10,11,12,13,14,15,16,17], [1,2,3,6,7,8]);
}

function groupVatDetailsForVatCredit_v621_(details, keyFn, initFn) {
  var map = {};
  (details || []).forEach(function(d) {
    var key = keyFn(d);
    if (!map[key]) {
      map[key] = initFn(d);
      map[key].qty = 0;
      map[key].netSales = 0;
      map[key].purchase = 0;
      map[key].profit = 0;
      map[key].rows = [];
    }
    var g = map[key];
    g.qty += d.qty || 0;
    g.netSales += d.netSales || 0;
    g.purchase += d.purchase || 0;
    g.profit += (d.estimatedProfit != null ? d.estimatedProfit : (d.netProfit || 0));
    g.rows.push(d);
  });
  return map;
}

function splitVatIncluded_v621_(amount) {
  amount = Math.round(toNumber_(amount));
  var supply = Math.round(amount / 1.1);
  return { total: amount, supply: supply, vat: amount - supply };
}

function vatPartFromVatIncluded_v621_(amount) {
  return splitVatIncluded_v621_(amount).vat;
}
