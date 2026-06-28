/**
 * LOTTEON v6.27 AC column purchase source-of-truth patch
 *
 * 확정 기준:
 * - 매출데이터_붙여넣기 시트의 AC열을 매입금액 원천 기준으로 사용합니다.
 * - 부가세_신고자료 / 부가세_상품별 / 부가세_고객별 / 부가세_주문번호별의 매입금액 합계는
 *   같은 집계 대상의 AC열 합계와 일치해야 합니다.
 * - 부가세 시트는 취소/반품 제외 순수매출 기준이므로, 검증도 취소/반품 제외 AC 합계와 비교합니다.
 *
 * 수정:
 * - 매출데이터_붙여넣기의 AC열 값을 행 순서 기준으로 detailRows에 직접 반영합니다.
 * - 행 수가 맞지 않을 때는 주문번호/상품번호/상품명/순수매출액 키 매칭으로 보강합니다.
 * - 반영 후 브랜드/월/전체 매입금액과 예상이익을 재계산합니다.
 * - 매입금액_AC검증 시트를 생성해 AC 합계와 부가세 상세 매입금액 합계 차이를 보여줍니다.
 */

var LOTTEON_PATCH_V627_AC_COLUMN_PURCHASE_SOURCE_OF_TRUTH_LOADED = true;
var LOTTEON_PURCHASE_AC_COLUMN_INDEX_V627 = 29; // AC column, 1-based

var __baseAggregateSalesByBrand_v627 = typeof aggregateSalesByBrand_v611_ === 'function' ? aggregateSalesByBrand_v611_ : null;
var __baseBuildVatBreakdownSheetsVatCredit_v627 = typeof buildVatBreakdownSheetsVatCredit_v621_ === 'function' ? buildVatBreakdownSheetsVatCredit_v621_ : null;
var __baseBuildBrandMarginSheet_v627 = typeof buildBrandMarginSheet_v611_ === 'function' ? buildBrandMarginSheet_v611_ : null;

aggregateSalesByBrand_v611_ = function() {
  var salesAgg = __baseAggregateSalesByBrand_v627 ? __baseAggregateSalesByBrand_v627.apply(this, arguments) : null;
  return applyAcColumnPurchaseSourceOfTruth_v627_(salesAgg);
};

buildVatBreakdownSheetsVatCredit_v621_ = function(salesAgg) {
  salesAgg = applyAcColumnPurchaseSourceOfTruth_v627_(salesAgg || aggregateSalesByBrand_v611_());
  return __baseBuildVatBreakdownSheetsVatCredit_v627 ? __baseBuildVatBreakdownSheetsVatCredit_v627.call(this, salesAgg) : null;
};

buildBrandMarginSheet_v611_ = function(salesAgg) {
  salesAgg = applyAcColumnPurchaseSourceOfTruth_v627_(salesAgg || aggregateSalesByBrand_v611_());
  return __baseBuildBrandMarginSheet_v627 ? __baseBuildBrandMarginSheet_v627.call(this, salesAgg) : null;
};

function applyAcColumnPurchaseSourceOfTruth_v627_(salesAgg) {
  salesAgg = salesAgg || {};
  if (salesAgg.__v627AcPurchaseApplied) return salesAgg;
  salesAgg.__v627AcPurchaseApplied = true;

  var details = salesAgg.detailRows || [];
  if (!details.length) return salesAgg;

  var ac = buildAcPurchaseRows_v627_();
  var matched = 0;
  var unmatched = [];
  var methodCounts = {};

  if (ac.netRows.length === details.length) {
    for (var i = 0; i < details.length; i++) {
      setDetailPurchaseFromAc_v627_(details[i], ac.netRows[i], 'AC행순서');
      matched++;
      methodCounts['AC행순서'] = (methodCounts['AC행순서'] || 0) + 1;
    }
  } else {
    var index = buildAcQueueIndex_v627_(ac.netRows);
    details.forEach(function(d, idx) {
      var picked = pickAcRowForDetail_v627_(index, d);
      if (picked) {
        setDetailPurchaseFromAc_v627_(d, picked.row, picked.method);
        matched++;
        methodCounts[picked.method] = (methodCounts[picked.method] || 0) + 1;
      } else {
        unmatched.push([idx + 1, d.orderNo || '', d.productNo || '', d.brand || '', compactTextForDiag_v626_(d.productName || '', 80), d.netSales || 0, d.purchase || 0, 'AC매칭실패']);
      }
    });
  }

  recomputeSalesAggPurchaseProfitFromDetails_v627_(salesAgg);
  writeAcPurchaseValidation_v627_(salesAgg, ac, matched, unmatched, methodCounts);

  try {
    log_('ac_purchase_source_v627', 'acNet=' + ac.netTotal + ' / detailPurchase=' + salesAgg.totalPurchase + ' / diff=' + (ac.netTotal - salesAgg.totalPurchase) + ' / matched=' + matched + ' / detailRows=' + details.length + ' / acRows=' + ac.netRows.length);
  } catch (e) {}

  return salesAgg;
}

function buildAcPurchaseRows_v627_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SALES_IN) || ss.getSheetByName('매출데이터_붙여넣기');
  var result = { rows: [], netRows: [], allTotal: 0, netTotal: 0, sheetName: sheet ? sheet.getName() : '', acHeader: '', rowCount: 0 };
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < LOTTEON_PURCHASE_AC_COLUMN_INDEX_V627) return result;

  var values = sheet.getDataRange().getValues();
  var header = values[0].map(function(h) { return String(h || '').trim(); });
  var acIdx = LOTTEON_PURCHASE_AC_COLUMN_INDEX_V627 - 1;
  result.acHeader = header[acIdx] || 'AC열';

  for (var r = 1; r < values.length; r++) {
    var obj = {};
    for (var c = 0; c < header.length; c++) if (header[c]) obj[header[c]] = values[r][c];
    var purchase = toNumber_(values[r][acIdx]);
    var salesAmount = toNumber_(getFirstValue_v613_(obj, ['순수매출액','결제금액합계','결제금액','매출액','판매금액','총결제금액','상품금액']));
    var orderNo = stringKey_v626_(getFirstValue_v613_(obj, ['마켓주문번호','주문번호','주문ID','주문번호_원본','주문 번호','마켓 주문번호']));
    var productNo = stringKey_v626_(getFirstValue_v613_(obj, ['사이트상품번호','마켓상품번호','상품번호','상품ID','상품코드','상품 번호','사이트 상품번호']));
    var productName = stringKey_v626_(getFirstValue_v613_(obj, ['마켓상품명','원문상품명','상품명','상품명_원본','옵션명','상품 명']));
    var cancel = typeof isCancelRow_v613_ === 'function' ? isCancelRow_v613_(obj) : false;
    var row = {
      sheetRow: r + 1,
      purchase: purchase,
      sales: salesAmount,
      orderNo: orderNo,
      productNo: productNo,
      productName: productName,
      nameKey: normalizeTextKey_v614_(productName),
      salesKey: moneyKey_v626_(salesAmount),
      cancel: cancel
    };
    result.rows.push(row);
    result.allTotal += purchase;
    if (!cancel) {
      result.netRows.push(row);
      result.netTotal += purchase;
    }
  }
  result.rowCount = result.rows.length;
  return result;
}

function buildAcQueueIndex_v627_(rows) {
  var index = { orderProductSales: {}, orderNameSales: {}, orderProduct: {}, orderName: {}, productSales: {}, nameSales: {} };
  (rows || []).forEach(function(row) {
    if (row.orderNo && row.productNo && row.sales) pushAcQueue_v627_(index.orderProductSales, row.orderNo + '||' + row.productNo + '||' + row.salesKey, row);
    if (row.orderNo && row.nameKey && row.sales) pushAcQueue_v627_(index.orderNameSales, row.orderNo + '||' + row.nameKey + '||' + row.salesKey, row);
    if (row.orderNo && row.productNo) pushAcQueue_v627_(index.orderProduct, row.orderNo + '||' + row.productNo, row);
    if (row.orderNo && row.nameKey) pushAcQueue_v627_(index.orderName, row.orderNo + '||' + row.nameKey, row);
    if (row.productNo && row.sales) pushAcQueue_v627_(index.productSales, row.productNo + '||' + row.salesKey, row);
    if (row.nameKey && row.sales) pushAcQueue_v627_(index.nameSales, row.nameKey + '||' + row.salesKey, row);
  });
  return index;
}

function pushAcQueue_v627_(map, key, row) {
  if (!map[key]) map[key] = [];
  map[key].push(row);
}

function popAcQueue_v627_(map, key) {
  var q = map && map[key];
  if (!q || !q.length) return null;
  return q.shift();
}

function pickAcRowForDetail_v627_(index, d) {
  var orderNo = stringKey_v626_(d.orderNo);
  var productNo = stringKey_v626_(d.productNo);
  var nameKey = normalizeTextKey_v614_(d.productName || '');
  var sales = toNumber_(d.netSales || 0);
  var salesKey = moneyKey_v626_(sales);
  var picked = null;

  if (!picked && orderNo && productNo && sales) picked = popAcQueue_v627_(index.orderProductSales, orderNo + '||' + productNo + '||' + salesKey);
  if (picked) return { row: picked, method: 'AC_주문+상품번호+매출액' };
  if (!picked && orderNo && nameKey && sales) picked = popAcQueue_v627_(index.orderNameSales, orderNo + '||' + nameKey + '||' + salesKey);
  if (picked) return { row: picked, method: 'AC_주문+상품명+매출액' };
  if (!picked && orderNo && productNo) picked = popAcQueue_v627_(index.orderProduct, orderNo + '||' + productNo);
  if (picked) return { row: picked, method: 'AC_주문+상품번호' };
  if (!picked && orderNo && nameKey) picked = popAcQueue_v627_(index.orderName, orderNo + '||' + nameKey);
  if (picked) return { row: picked, method: 'AC_주문+상품명' };
  if (!picked && productNo && sales) picked = popAcQueue_v627_(index.productSales, productNo + '||' + salesKey);
  if (picked) return { row: picked, method: 'AC_상품번호+매출액' };
  if (!picked && nameKey && sales) picked = popAcQueue_v627_(index.nameSales, nameKey + '||' + salesKey);
  if (picked) return { row: picked, method: 'AC_상품명+매출액' };
  return null;
}

function setDetailPurchaseFromAc_v627_(d, row, method) {
  d.purchase = toNumber_(row.purchase);
  d.purchaseMatchType = method;
  d.purchaseMatchKey = 'AC row ' + row.sheetRow;
  d.estimatedProfit = toNumber_(d.settlementBasis || d.settlement) - d.purchase;
  d.netProfit = d.estimatedProfit;
}

function recomputeSalesAggPurchaseProfitFromDetails_v627_(salesAgg) {
  if (typeof recomputeSalesAggPurchaseProfit_v626_ === 'function') {
    recomputeSalesAggPurchaseProfit_v626_(salesAgg);
    return;
  }
}

function writeAcPurchaseValidation_v627_(salesAgg, ac, matched, unmatched, methodCounts) {
  var ss = SpreadsheetApp.getActive();
  var totalPurchase = toNumber_(salesAgg.totalPurchase || 0);
  var diff = Math.round(ac.netTotal - totalPurchase);
  var sheet = ss.getSheetByName('매입금액_AC검증') || ss.insertSheet('매입금액_AC검증');
  var rows = [
    ['항목','값','메모'],
    ['원본시트', ac.sheetName, 'AC열 기준'],
    ['AC열 헤더', ac.acHeader, '매출데이터_붙여넣기 AC열'],
    ['AC 전체 합계', ac.allTotal, '취소/반품 포함'],
    ['AC 취소제외 합계', ac.netTotal, '부가세 시트 비교 기준'],
    ['부가세 상세 매입금액 합계', totalPurchase, 'detailRows purchase 합계'],
    ['차이', diff, 'AC 취소제외 합계 - 부가세 상세 매입금액 합계'],
    ['AC 취소제외 행수', ac.netRows.length, ''],
    ['부가세 상세 행수', (salesAgg.detailRows || []).length, ''],
    ['매칭 행수', matched, ''],
    ['미매칭 행수', unmatched.length, ''],
    ['매칭방식', JSON.stringify(methodCounts || {}), '']
  ];
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  try {
    sheet.getRange(1, 1, 1, 3).setBackground('#d9eaf7').setFontWeight('bold');
    sheet.getRange(4, 2, 3, 1).setNumberFormat('#,##0');
    sheet.autoResizeColumns(1, 3);
  } catch (e) {}

  var miss = ss.getSheetByName('매입금액_AC미매칭') || ss.insertSheet('매입금액_AC미매칭');
  var headers = ['행번호','주문번호','상품번호','브랜드명','상품명','순수매출액','기존매입금액','사유'];
  replaceSheetValues_v611_(miss, headers, unmatched.slice(0, 500));
  try { formatMoneyColumns_v611_(miss, [6,7]); } catch (e) {}
}
