/**
 * LOTTEON v6.28 single-source financial validation patch
 *
 * 확정 기준:
 * - 매출데이터_붙여넣기 시트를 모든 매출/매입/정산/부가세 집계의 단일 원천으로 사용합니다.
 * - 분석대상 행 기준:
 *   1) 취소/반품/교환 상태 제외
 *   2) 구매사이트명에 lotteon 포함
 *   3) 사이트상품번호 있음
 *   4) 브랜드 있음
 * - 매출금액 원천: 결제금액합계(원)
 * - 매입금액 원천: AC열 구매가격
 * - 실제정산금액 원천: 정산예정금액(원)
 * - 정산예정금액이 비어 있거나 0이면 예상정산금액 = 순수매출액 × 0.901
 *
 * 수정:
 * - aggregateSalesByBrand_v611_를 원본 단일 원천 집계로 대체합니다.
 * - 대시보드 / 브랜드별_마진율 / 부가세_신고자료 / 부가세_상품별 / 부가세_고객별 / 부가세_주문번호별을 같은 detailRows 기준으로 재생성합니다.
 * - 시트별_금액검증 / 원본대상행_검증 / 분석제외_행_리스트를 생성합니다.
 */

var LOTTEON_PATCH_V628_SINGLE_SOURCE_FINANCIAL_VALIDATION_LOADED = true;
var LOTTEON_V628_SETTLEMENT_RATE = 0.901;
var LOTTEON_V628_PURCHASE_AC_COL = 29; // AC, 1-based

aggregateSalesByBrand_v611_ = function() {
  return buildSingleSourceSalesAgg_v628_();
};

rebuildDashboardBrandBased_v611_ = function(salesAgg, filterAgg) {
  return rebuildDashboardSingleSource_v628_(salesAgg || buildSingleSourceSalesAgg_v628_(), filterAgg || (typeof aggregateFiltersByBrand_v611_ === 'function' ? aggregateFiltersByBrand_v611_() : {}));
};

buildBrandMarginSheet_v611_ = function(salesAgg) {
  return buildBrandMarginSingleSource_v628_(salesAgg || buildSingleSourceSalesAgg_v628_());
};

buildVatBreakdownSheetsVatCredit_v621_ = function(salesAgg) {
  return buildVatReportsSingleSource_v628_(salesAgg || buildSingleSourceSalesAgg_v628_());
};

buildVatReportSheet_v611_ = function(salesAgg) {
  return buildVatReportsSingleSource_v628_(salesAgg || buildSingleSourceSalesAgg_v628_());
};

refreshDashboardFastOnly = function() {
  var ui = SpreadsheetApp.getUi();
  var started = Date.now();
  try {
    var salesAgg = buildSingleSourceSalesAgg_v628_();
    var filterAgg = typeof aggregateFiltersByBrand_v611_ === 'function' ? aggregateFiltersByBrand_v611_() : {};
    rebuildDashboardSingleSource_v628_(salesAgg, filterAgg);
    buildBrandMarginSingleSource_v628_(salesAgg);
    if (typeof buildUnsettledSettlementByAccountSheet_v616_ === 'function') buildUnsettledSettlementByAccountSheet_v616_(salesAgg);
    buildFinancialValidationSheets_v628_(salesAgg);
    if (typeof applyFastOutputSheetFormatting_v620_ === 'function') applyFastOutputSheetFormatting_v620_();
    ui.alert('대시보드 빠른 갱신 완료\n\n기준: v6.28 매출데이터_붙여넣기 단일 원천\n소요초: ' + Math.round((Date.now() - started) / 1000));
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    try { log_('dashboard_fast_v628_error', msg); } catch (ignore) {}
    ui.alert('대시보드 빠른 갱신 중 오류가 발생했습니다.\n\n' + msg);
    throw e;
  }
};

function generateVatReportsFullSeparated_v622() {
  var ui = SpreadsheetApp.getUi();
  var started = Date.now();
  try {
    var salesAgg = buildSingleSourceSalesAgg_v628_();
    buildVatReportsSingleSource_v628_(salesAgg);
    buildBrandMarginSingleSource_v628_(salesAgg);
    rebuildDashboardSingleSource_v628_(salesAgg, typeof aggregateFiltersByBrand_v611_ === 'function' ? aggregateFiltersByBrand_v611_() : {});
    buildFinancialValidationSheets_v628_(salesAgg);
    if (typeof applyFastOutputSheetFormatting_v620_ === 'function') applyFastOutputSheetFormatting_v620_();
    ui.alert(
      '부가세 신고자료 생성 완료\n\n' +
      '기준: v6.28 매출데이터_붙여넣기 단일 원천\n' +
      '분석대상 행수: ' + salesAgg.detailRows.length + '\n' +
      '순수매출액: ' + formatNumberForMsg_v628_(salesAgg.totalNetSales) + '\n' +
      '매입금액: ' + formatNumberForMsg_v628_(salesAgg.totalPurchase) + '\n' +
      '소요초: ' + Math.round((Date.now() - started) / 1000) + '\n\n' +
      '시트별_금액검증에서 차이가 0인지 확인하세요.'
    );
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    try { log_('vat_reports_v628_error', msg); } catch (ignore) {}
    ui.alert('부가세 신고자료 생성 중 오류가 발생했습니다.\n\n' + msg);
    throw e;
  }
}

function buildSingleSourceSalesAgg_v628_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SALES_IN) || ss.getSheetByName('매출데이터_붙여넣기');
  if (!sheet || sheet.getLastRow() < 2) throw new Error('매출데이터_붙여넣기 시트를 찾지 못했거나 데이터가 없습니다.');

  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function(h) { return String(h || '').trim(); });
  var acIdx = LOTTEON_V628_PURCHASE_AC_COL - 1;

  var agg = initSalesAgg_v628_(sheet.getName(), headers[acIdx] || 'AC열');
  var targetRows = [];
  var excludedRows = [];

  for (var r = 1; r < values.length; r++) {
    var rowObj = rowToObject_v628_(headers, values[r]);
    var sales = num_v628_(getAny_v628_(rowObj, ['결제금액합계(원)','결제금액합계','결제금액','매출액','판매금액']));
    var purchase = num_v628_(values[r][acIdx]);
    var row = makeSourceRow_v628_(r + 1, rowObj, sales, purchase);

    agg.rawRowCount++;
    agg.totalGrossSales += sales;
    agg.totalRawPurchase += purchase;

    var exclude = getAnalysisExclusionReason_v628_(row);
    if (exclude) {
      row.excludeReason = exclude;
      excludedRows.push(row);
      if (row.cancelExcluded) {
        agg.cancelRowCount++;
        agg.totalCancelSales += sales;
        agg.totalCancelPurchase += purchase;
      } else {
        agg.analysisExcludedRowCount++;
        agg.totalAnalysisExcludedSales += sales;
        agg.totalAnalysisExcludedPurchase += purchase;
      }
      continue;
    }

    targetRows.push(row);
  }

  agg.excludedRows = excludedRows;
  agg.targetSourceRows = targetRows;

  targetRows.forEach(function(row) {
    addDetailRowToAgg_v628_(agg, sourceRowToDetail_v628_(row));
  });

  finalizeSalesAgg_v628_(agg);
  try {
    log_('single_source_agg_v628', 'raw=' + agg.rawRowCount + ' target=' + agg.detailRows.length + ' net=' + agg.totalNetSales + ' purchase=' + agg.totalPurchase + ' excluded=' + agg.excludedRows.length);
  } catch (e) {}
  return agg;
}

function initSalesAgg_v628_(sourceSheetName, acHeader) {
  return {
    sourceSheetName: sourceSheetName,
    acHeader: acHeader,
    basisVersion: 'v6.28 single source',
    byBrand: {},
    detailRows: [],
    excludedRows: [],
    targetSourceRows: [],
    rawRowCount: 0,
    cancelRowCount: 0,
    analysisExcludedRowCount: 0,
    totalGrossSales: 0,
    totalCancelSales: 0,
    totalAnalysisExcludedSales: 0,
    totalExcludedSales: 0,
    totalNetSales: 0,
    totalRawPurchase: 0,
    totalCancelPurchase: 0,
    totalAnalysisExcludedPurchase: 0,
    totalPurchase: 0,
    totalActualSettlement: 0,
    totalEstimatedSettlement: 0,
    totalSettlementBasis: 0,
    totalMarketFee: 0,
    totalEstimatedProfit: 0,
    totalNetProfit: 0,
    totalOrders: 0,
    totalProducts: 0,
    totalQuantity: 0,
    totalMarketFeeRate: 0,
    totalEstimatedProfitRate: 0,
    totalNetProfitRate: 0,
    hasPurchaseColumn: true,
    productSet: {},
    orderSet: {},
    unsettledOrderSet: {},
    overdueUnsettledOrderSet: {}
  };
}

function makeSourceRow_v628_(sheetRow, obj, sales, purchase) {
  var orderNo = normalizeId_v628_(getAny_v628_(obj, ['마켓주문번호','주문번호','주문ID','주문 번호','마켓 주문번호']));
  var productNo = normalizeId_v628_(getAny_v628_(obj, ['사이트상품번호','마켓상품번호','상품번호','상품ID','상품코드','상품 번호','사이트 상품번호']));
  var marketProductNo = normalizeId_v628_(getAny_v628_(obj, ['마켓상품번호','상품번호','마켓 상품번호']));
  var brand = String(getAny_v628_(obj, ['브랜드','브랜드명']) || '').trim();
  var productName = String(getAny_v628_(obj, ['마켓상품명','상품명(원문)','상품명','원문상품명']) || '').trim();
  var purchaseSite = String(getAny_v628_(obj, ['구매사이트명','구매 사이트명']) || '').trim();
  var marketStatus = String(getAny_v628_(obj, ['마켓주문상태','주문상태']) || '').trim();
  var themangoStatus = String(getAny_v628_(obj, ['더망고주문상태','더망고 주문상태']) || '').trim();
  var qty = num_v628_(getAny_v628_(obj, ['결제수량','판매수량','수량'])) || 0;
  var actualSettlement = num_v628_(getAny_v628_(obj, ['정산예정금액(원)','정산예정금액','실제정산금액']));
  var dateObj = parseDateValue_v628_(getAny_v628_(obj, ['마켓주문일자','주문일자','주문일']));
  var cancel = isCancelStatus_v628_(marketStatus, themangoStatus);

  return {
    sheetRow: sheetRow,
    obj: obj,
    orderNo: orderNo,
    productNo: productNo,
    marketProductNo: marketProductNo,
    brand: brand,
    productName: productName,
    purchaseSite: purchaseSite,
    marketStatus: marketStatus,
    themangoStatus: themangoStatus,
    qty: qty,
    sales: sales,
    purchase: purchase,
    actualSettlement: actualSettlement,
    date: dateObj,
    dateText: formatDateMMDD_v628_(dateObj),
    year: dateObj ? Number(Utilities.formatDate(dateObj, CONFIG.TIMEZONE, 'yyyy')) : '',
    month: dateObj ? Number(Utilities.formatDate(dateObj, CONFIG.TIMEZONE, 'M')) : '',
    day: dateObj ? Number(Utilities.formatDate(dateObj, CONFIG.TIMEZONE, 'd')) : '',
    cancelExcluded: cancel
  };
}

function getAnalysisExclusionReason_v628_(row) {
  if (row.cancelExcluded) return '취소/반품/교환 상태 제외';
  if (!/lotteon/i.test(row.purchaseSite || '')) return '구매사이트 lotteon 아님/공란';
  if (!row.productNo) return '사이트상품번호 없음';
  if (!row.brand) return '브랜드 없음';
  return '';
}

function sourceRowToDetail_v628_(row) {
  var estimatedSettlement = row.actualSettlement ? 0 : Math.round(row.sales * LOTTEON_V628_SETTLEMENT_RATE);
  var settlementBasis = row.actualSettlement || estimatedSettlement;
  var marketFee = row.sales - settlementBasis;
  var profit = settlementBasis - row.purchase;
  var orderDate = row.date;
  var overdue = false;
  if (!row.actualSettlement && orderDate) {
    overdue = ((new Date()).getTime() - orderDate.getTime()) / (1000 * 60 * 60 * 24) > 30;
  }
  return {
    sourceRow: row.sheetRow,
    dateText: row.dateText,
    year: row.year,
    month: row.month,
    day: row.day,
    orderDate: orderDate,
    orderNo: row.orderNo,
    customer: '',
    brand: row.brand,
    productNo: row.productNo,
    marketProductNo: row.marketProductNo,
    productName: row.productName,
    qty: row.qty,
    grossSales: row.sales,
    netSales: row.sales,
    salesAmount: row.sales,
    actualSettlement: row.actualSettlement,
    estimatedSettlement: estimatedSettlement,
    settlementBasis: settlementBasis,
    settlement: settlementBasis,
    marketFee: marketFee,
    purchase: row.purchase,
    purchaseAmount: row.purchase,
    estimatedProfit: profit,
    netProfit: profit,
    note: row.actualSettlement ? '정산완료' : '미정산',
    unsettledStatus: row.actualSettlement ? '' : '미정산',
    overdueStatus: overdue ? '30일초과 미정산' : '',
    purchaseMatchType: 'AC열 직접',
    purchaseMatchKey: 'source row ' + row.sheetRow
  };
}

function addDetailRowToAgg_v628_(agg, d) {
  agg.detailRows.push(d);
  agg.totalNetSales += d.netSales;
  agg.totalPurchase += d.purchase;
  agg.totalActualSettlement += d.actualSettlement;
  agg.totalEstimatedSettlement += d.estimatedSettlement;
  agg.totalSettlementBasis += d.settlementBasis;
  agg.totalMarketFee += d.marketFee;
  agg.totalEstimatedProfit += d.estimatedProfit;
  agg.totalNetProfit += d.netProfit;
  agg.totalQuantity += d.qty;
  if (d.orderNo) agg.orderSet[d.orderNo] = true;
  var productKey = (d.productNo || '') + '|' + (d.productName || '');
  if (productKey !== '|') agg.productSet[productKey] = true;
  if (d.unsettledStatus && d.orderNo) agg.unsettledOrderSet[d.orderNo] = true;
  if (d.overdueStatus && d.orderNo) agg.overdueUnsettledOrderSet[d.orderNo] = true;

  var brandKey = brandKey_v628_(d.brand);
  if (!agg.byBrand[brandKey]) agg.byBrand[brandKey] = initBrandAgg_v628_(d.brand);
  addDetailToBrandAgg_v628_(agg.byBrand[brandKey], d);
}

function initBrandAgg_v628_(brand) {
  return {
    brand: brand || '브랜드미확인',
    orderSet: {}, productSet: {}, customerSet: {}, months: {},
    orderCount: 0, customerCount: 0, quantity: 0, salesProductCount: 0,
    grossSalesAmount: 0, cancelSalesAmount: 0, netSalesAmount: 0, salesAmount: 0,
    actualSettlementAmount: 0, estimatedSettlementAmount: 0, settlementBasisAmount: 0, settlementAmount: 0,
    marketFeeAmount: 0, marketFeeRate: 0, purchaseAmount: 0,
    estimatedProfitAmount: 0, netProfitAmount: 0, marginAmount: 0,
    estimatedProfitRate: 0, netProfitRate: 0, marginRate: 0,
    unsettledOrderSet: {}, overdueUnsettledOrderSet: {}, unsettledOrderCount: 0, overdueUnsettledOrderCount: 0,
    unsettledNetSales: 0
  };
}

function addDetailToBrandAgg_v628_(b, d) {
  if (d.orderNo) b.orderSet[d.orderNo] = true;
  var productKey = (d.productNo || '') + '|' + (d.productName || '');
  if (productKey !== '|') b.productSet[productKey] = true;
  b.quantity += d.qty || 0;
  b.grossSalesAmount += d.netSales || 0;
  b.netSalesAmount += d.netSales || 0;
  b.salesAmount += d.netSales || 0;
  b.actualSettlementAmount += d.actualSettlement || 0;
  b.estimatedSettlementAmount += d.estimatedSettlement || 0;
  b.settlementBasisAmount += d.settlementBasis || 0;
  b.settlementAmount += d.settlementBasis || 0;
  b.marketFeeAmount += d.marketFee || 0;
  b.purchaseAmount += d.purchase || 0;
  b.estimatedProfitAmount += d.estimatedProfit || 0;
  b.netProfitAmount += d.netProfit || 0;
  b.marginAmount += d.netProfit || 0;
  if (d.unsettledStatus && d.orderNo) { b.unsettledOrderSet[d.orderNo] = true; b.unsettledNetSales += d.netSales || 0; }
  if (d.overdueStatus && d.orderNo) b.overdueUnsettledOrderSet[d.orderNo] = true;

  var monthKey = d.year && d.month ? d.year + '-' + ('0' + d.month).slice(-2) : 'unknown';
  if (!b.months[monthKey]) b.months[monthKey] = { sales: 0, purchase: 0, quantity: 0, orders: {}, products: {}, estimatedProfitAmount: 0, netProfitAmount: 0, netProfitRate: 0 };
  b.months[monthKey].sales += d.netSales || 0;
  b.months[monthKey].purchase += d.purchase || 0;
  b.months[monthKey].quantity += d.qty || 0;
  if (d.orderNo) b.months[monthKey].orders[d.orderNo] = true;
  b.months[monthKey].products[productKey] = true;
  b.months[monthKey].estimatedProfitAmount += d.estimatedProfit || 0;
  b.months[monthKey].netProfitAmount += d.netProfit || 0;
}

function finalizeSalesAgg_v628_(agg) {
  agg.totalExcludedSales = agg.totalCancelSales + agg.totalAnalysisExcludedSales;
  agg.totalExcludedPurchase = agg.totalCancelPurchase + agg.totalAnalysisExcludedPurchase;
  agg.totalOrders = Object.keys(agg.orderSet).length;
  agg.totalProducts = Object.keys(agg.productSet).length;
  agg.unsettledOrderCount = Object.keys(agg.unsettledOrderSet).length;
  agg.overdueUnsettledOrderCount = Object.keys(agg.overdueUnsettledOrderSet).length;
  agg.totalMarketFeeRate = agg.totalNetSales ? agg.totalMarketFee / agg.totalNetSales : 0;
  agg.totalEstimatedProfitRate = agg.totalNetSales ? agg.totalEstimatedProfit / agg.totalNetSales : 0;
  agg.totalNetProfitRate = agg.totalEstimatedProfitRate;
  Object.keys(agg.byBrand).forEach(function(k) {
    var b = agg.byBrand[k];
    b.orderCount = Object.keys(b.orderSet).length;
    b.salesProductCount = Object.keys(b.productSet).length;
    b.customerCount = Object.keys(b.customerSet).length;
    b.unsettledOrderCount = Object.keys(b.unsettledOrderSet).length;
    b.overdueUnsettledOrderCount = Object.keys(b.overdueUnsettledOrderSet).length;
    b.marketFeeRate = b.netSalesAmount ? b.marketFeeAmount / b.netSalesAmount : 0;
    b.estimatedProfitRate = b.netSalesAmount ? b.estimatedProfitAmount / b.netSalesAmount : 0;
    b.netProfitRate = b.estimatedProfitRate;
    b.marginRate = b.estimatedProfitRate;
    Object.keys(b.months || {}).forEach(function(mk) {
      var m = b.months[mk];
      m.orderCount = Object.keys(m.orders || {}).length;
      m.productCount = Object.keys(m.products || {}).length;
      m.netProfitRate = m.sales ? m.netProfitAmount / m.sales : 0;
    });
  });
}

function rebuildDashboardSingleSource_v628_(salesAgg, filterAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD) || ss.insertSheet(CONFIG.SHEETS.DASHBOARD);
  var rows = [];
  var nowText = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  var supply = Math.round(salesAgg.totalNetSales / 1.1);
  var vat = salesAgg.totalNetSales - supply;
  var purchaseVat = salesAgg.totalPurchase - Math.round(salesAgg.totalPurchase / 1.1);
  var payableVat = vat - purchaseVat;

  rows.push(['요약','갱신기준','v6.28 매출데이터_붙여넣기 단일 원천','','','','','','','','','','','','','']);
  rows.push(['요약','갱신시각',nowText,'','','','','','','','','','','','','Asia/Seoul']);
  rows.push(['요약','원본 전체 행수',salesAgg.rawRowCount,'','','','','','','','','','','','','매출데이터_붙여넣기']);
  rows.push(['요약','분석대상 행수',salesAgg.detailRows.length,'','','','','','','','','','','','','취소/반품 제외 + lotteon + 상품번호 + 브랜드']);
  rows.push(['요약','분석브랜드수',Object.keys(salesAgg.byBrand).length,'','','','','','','','','','','','','브랜드 기준']);
  rows.push(['요약','수집상품수',filterAgg && filterAgg.totalCollectCount || 0,'','','','','','','','','','','','','필터별_상품수 API_totalCount 총합']);
  rows.push(['요약','매출상품수',salesAgg.totalProducts,'','','','','','','','','','','','','분석대상 상품수']);
  rows.push(['요약','주문건수',salesAgg.totalOrders,'','','','','','','','','','','','','분석대상 주문수']);
  rows.push(['요약','미정산 주문건수',salesAgg.unsettledOrderCount,'','','','','','','','','','','','','정산예정금액 공란/0']);
  rows.push(['요약','30일초과 미정산 주문건수',salesAgg.overdueUnsettledOrderCount,'','','','','','','','','','','','','주문일 기준 30일 초과']);
  rows.push(['요약','총매출액',salesAgg.totalGrossSales,'','','','','','','','','','','','','원본 전체']);
  rows.push(['요약','취소/반품매출',salesAgg.totalCancelSales,'','','','','','','','','','','','','상태값 제외']);
  rows.push(['요약','분석제외매출',salesAgg.totalAnalysisExcludedSales,'','','','','','','','','','','','','취소 외 분석대상 제외']);
  rows.push(['요약','제외매출합계',salesAgg.totalExcludedSales,'','','','','','','','','','','','','취소/반품 + 분석제외']);
  rows.push(['요약','순수매출액',salesAgg.totalNetSales,'','','','','','','','','','','','','분석대상 기준']);
  rows.push(['요약','실제정산금액',salesAgg.totalActualSettlement,'','','','','','','','','','','','','정산예정금액 입력 완료']);
  rows.push(['요약','예상정산금액(미정산)',salesAgg.totalEstimatedSettlement,'','','','','','','','','','','','','미정산 순수매출액 × 0.901']);
  rows.push(['요약','정산기준금액',salesAgg.totalSettlementBasis,'','','','','','','','','','','','','실제 + 예상']);
  rows.push(['요약','마켓수수료/비용',salesAgg.totalMarketFee,'','','','','','','','','','','','','순수매출액 - 정산기준금액']);
  rows.push(['요약','마켓수수료율',salesAgg.totalMarketFeeRate,'','','','','','','','','','','','','마켓수수료 ÷ 순수매출액']);
  rows.push(['요약','매입금액',salesAgg.totalPurchase,'','','','','','','','','','','','','AC열 구매가격 직접 합계']);
  rows.push(['요약','예상이익',salesAgg.totalEstimatedProfit,'','','','','','','','','','','','','정산기준금액 - 매입금액']);
  rows.push(['요약','예상이익률',salesAgg.totalEstimatedProfitRate,'','','','','','','','','','','','','예상이익 ÷ 순수매출액']);
  rows.push(['요약','부가세 신고 공급대가',salesAgg.totalNetSales,'','','','','','','','','','','','','순수매출액 기준']);
  rows.push(['요약','부가세 신고 공급가액',supply,'','','','','','','','','','','','','공급대가 ÷ 1.1']);
  rows.push(['요약','매출부가세',vat,'','','','','','','','','','','','','공급대가 - 공급가액']);
  rows.push(['요약','매입부가세',purchaseVat,'','','','','','','','','','','','','매입금액 부가세 포함 기준']);
  rows.push(['요약','납부예상부가세',payableVat,'','','','','','','','','','','','','매출부가세 - 매입부가세']);
  rows.push(['','','','','','','','','','','','','','','','']);

  var brandRows = buildDashboardBrandRows_v628_(salesAgg, filterAgg);
  rows.push(['구분','브랜드명','대표검색필터명','계정ID','수집수','매출상품','주문','순수매출액','실제정산금액','예상정산금액(미정산)','정산기준금액','매입금액','예상이익','예상이익률','미정산건수','30일초과건수']);
  brandRows.sort(function(a,b){return b.netSalesAmount - a.netSalesAmount;}).slice(0, 60).forEach(function(b) {
    rows.push(['브랜드TOP', b.brand, b.filterName, b.accountId, b.collectCount, b.salesProductCount, b.orderCount, b.netSalesAmount, b.actualSettlementAmount, b.estimatedSettlementAmount, b.settlementBasisAmount, b.purchaseAmount, b.estimatedProfitAmount, b.estimatedProfitRate, b.unsettledOrderCount, b.overdueUnsettledOrderCount]);
  });
  rows.push(['','','','','','','','','','','','','','','','']);
  rows.push(['구분','브랜드명','대표검색필터명','계정ID','수집수','매출상품','주문','순수매출액','정산기준금액','예상이익률','미정산건수','30일초과건수','다음작업','최근수집일','미판매수','메모']);
  brandRows.sort(function(a,b){return b.unsoldCount - a.unsoldCount;}).slice(0, 80).forEach(function(b) {
    rows.push(['상품갈이', b.brand, b.filterName, b.accountId, b.collectCount, b.salesProductCount, b.orderCount, b.netSalesAmount, b.settlementBasisAmount, b.estimatedProfitRate, b.unsettledOrderCount, b.overdueUnsettledOrderCount, decideNextRotationSafe_v628_(b), formatShortDateSafe_v628_(b.latestRecentDate), b.unsoldCount, 'v6.28 단일 원천']);
  });

  replaceSheetValuesSafe_v628_(sheet, ['구분','항목','값1','값2','값3','값4','값5','값6','값7','값8','값9','값10','값11','값12','값13','메모'], rows);
  formatOutputSheetBasic_v628_(sheet);
  return { rows: rows.length };
}

function buildBrandMarginSingleSource_v628_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('브랜드별_마진율') || ss.insertSheet('브랜드별_마진율');
  var headers = ['브랜드명','주문건수','고객수','판매수량','매출상품수','총매출액','취소/반품매출','순수매출액','실제정산금액','예상정산금액(미정산)','정산기준금액','마켓수수료/비용','마켓수수료율','매입금액','예상이익','예상이익률','매출부가세','매입부가세','납부예상부가세','부가세반영예상이익','부가세반영이익률','미정산주문건수','30일초과미정산건수','비고'];
  var rows = Object.keys(salesAgg.byBrand).map(function(k) {
    var b = salesAgg.byBrand[k];
    var salesVat = vatPart_v628_(b.netSalesAmount);
    var purchaseVat = vatPart_v628_(b.purchaseAmount);
    var payableVat = salesVat - purchaseVat;
    var afterVatProfit = b.estimatedProfitAmount - payableVat;
    return [b.brand, b.orderCount, b.customerCount, b.quantity, b.salesProductCount, b.grossSalesAmount, b.cancelSalesAmount, b.netSalesAmount, b.actualSettlementAmount, b.estimatedSettlementAmount, b.settlementBasisAmount, b.marketFeeAmount, b.marketFeeRate, b.purchaseAmount, b.estimatedProfitAmount, b.estimatedProfitRate, salesVat, purchaseVat, payableVat, afterVatProfit, b.netSalesAmount ? afterVatProfit / b.netSalesAmount : 0, b.unsettledOrderCount, b.overdueUnsettledOrderCount, 'v6.28 단일 원천'];
  }).sort(function(a,b){return b[7]-a[7];});
  replaceSheetValuesSafe_v628_(sheet, headers, rows);
  formatOutputSheetBasic_v628_(sheet);
  return { rows: rows.length };
}

function buildVatReportsSingleSource_v628_(salesAgg) {
  buildVatDetailSingleSource_v628_(salesAgg);
  buildVatProductSingleSource_v628_(salesAgg);
  buildVatCustomerSingleSource_v628_(salesAgg);
  buildVatOrderSingleSource_v628_(salesAgg);
  buildFinancialValidationSheets_v628_(salesAgg);
  return { rows: salesAgg.detailRows.length };
}

function buildVatDetailSingleSource_v628_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('부가세_신고자료') || ss.insertSheet('부가세_신고자료');
  var headers = ['날짜','주문번호','고객명','브랜드명','상품번호','상품명','판매수량','순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료/비용','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','비고'];
  var rows = salesAgg.detailRows.map(function(d) { return vatDetailRow_v628_(d); });
  replaceSheetValuesSafe_v628_(sheet, headers, rows);
  formatOutputSheetBasic_v628_(sheet);
}

function vatDetailRow_v628_(d) {
  var salesSplit = splitVat_v628_(d.netSales);
  var purchaseSplit = splitVat_v628_(d.purchase);
  var payableVat = salesSplit.vat - purchaseSplit.vat;
  return [d.dateText, d.orderNo, d.customer, d.brand, d.productNo, d.productName, d.qty, d.netSales, salesSplit.supply, salesSplit.vat, d.settlementBasis, d.marketFee, d.purchase, purchaseSplit.supply, purchaseSplit.vat, payableVat, d.estimatedProfit, d.estimatedProfit - payableVat, 'v6.28 AC열 직접'];
}

function buildVatProductSingleSource_v628_(salesAgg) {
  var map = groupDetails_v628_(salesAgg.detailRows, function(d){ return [d.brand,d.productNo,d.productName].join('|'); }, function(d){ return { brand:d.brand, productNo:d.productNo, productName:d.productName, orderSet:{}, customerSet:{} }; });
  var rows = Object.keys(map).map(function(k){ var g=map[k], ss=splitVat_v628_(g.netSales), ps=splitVat_v628_(g.purchase), pv=ss.vat-ps.vat; return [g.brand,g.productNo,g.productName,g.qty,g.netSales,ss.supply,ss.vat,g.purchase,ps.supply,ps.vat,pv,g.profit,g.profit-pv,Object.keys(g.orderSet).length,Object.keys(g.customerSet).length]; }).sort(function(a,b){return b[4]-a[4];});
  var sheet = SpreadsheetApp.getActive().getSheetByName('부가세_상품별') || SpreadsheetApp.getActive().insertSheet('부가세_상품별');
  replaceSheetValuesSafe_v628_(sheet, ['브랜드명','상품번호','상품명','판매수량','순수매출액','매출공급가액','매출부가세','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','주문건수','고객수'], rows);
  formatOutputSheetBasic_v628_(sheet);
}

function buildVatCustomerSingleSource_v628_(salesAgg) {
  var map = groupDetails_v628_(salesAgg.detailRows, function(d){ return d.customer || '고객명미확인'; }, function(d){ return { customer:d.customer || '고객명미확인', orderSet:{}, productSet:{} }; });
  var rows = Object.keys(map).map(function(k){ var g=map[k], ss=splitVat_v628_(g.netSales), ps=splitVat_v628_(g.purchase), pv=ss.vat-ps.vat; return [g.customer,Object.keys(g.orderSet).length,Object.keys(g.productSet).length,g.qty,g.netSales,ss.supply,ss.vat,g.purchase,ps.supply,ps.vat,pv,g.profit,g.profit-pv]; }).sort(function(a,b){return b[4]-a[4];});
  var sheet = SpreadsheetApp.getActive().getSheetByName('부가세_고객별') || SpreadsheetApp.getActive().insertSheet('부가세_고객별');
  replaceSheetValuesSafe_v628_(sheet, ['고객명','주문건수','상품수','판매수량','순수매출액','매출공급가액','매출부가세','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익'], rows);
  formatOutputSheetBasic_v628_(sheet);
}

function buildVatOrderSingleSource_v628_(salesAgg) {
  var map = groupDetails_v628_(salesAgg.detailRows, function(d){ return d.orderNo || '주문번호미확인'; }, function(d){ return { dateText:d.dateText, orderNo:d.orderNo || '주문번호미확인', customer:d.customer, brandSet:{}, productSet:{} }; });
  var rows = Object.keys(map).map(function(k){ var g=map[k], ss=splitVat_v628_(g.netSales), ps=splitVat_v628_(g.purchase), pv=ss.vat-ps.vat; return [g.dateText,g.orderNo,g.customer,Object.keys(g.brandSet).length,Object.keys(g.productSet).length,g.qty,g.netSales,ss.supply,ss.vat,g.purchase,ps.supply,ps.vat,pv,g.profit,g.profit-pv]; }).sort(function(a,b){ return String(a[0]).localeCompare(String(b[0])) || String(a[1]).localeCompare(String(b[1])); });
  var sheet = SpreadsheetApp.getActive().getSheetByName('부가세_주문번호별') || SpreadsheetApp.getActive().insertSheet('부가세_주문번호별');
  replaceSheetValuesSafe_v628_(sheet, ['날짜','주문번호','고객명','브랜드수','상품수','판매수량','순수매출액','매출공급가액','매출부가세','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익'], rows);
  formatOutputSheetBasic_v628_(sheet);
}

function groupDetails_v628_(details, keyFn, initFn) {
  var map = {};
  details.forEach(function(d) {
    var key = keyFn(d);
    if (!map[key]) { map[key] = initFn(d); map[key].qty=0; map[key].netSales=0; map[key].purchase=0; map[key].profit=0; }
    var g = map[key];
    g.qty += d.qty || 0;
    g.netSales += d.netSales || 0;
    g.purchase += d.purchase || 0;
    g.profit += d.estimatedProfit || 0;
    if (d.orderNo && g.orderSet) g.orderSet[d.orderNo] = true;
    if (d.customer && g.customerSet) g.customerSet[d.customer] = true;
    if ((d.productNo || d.productName) && g.productSet) g.productSet[(d.productNo || '') + '|' + (d.productName || '')] = true;
    if (d.brand && g.brandSet) g.brandSet[d.brand] = true;
  });
  return map;
}

function buildFinancialValidationSheets_v628_(salesAgg) {
  buildSourceBasisValidation_v628_(salesAgg);
  buildExcludedRowsSheet_v628_(salesAgg);
  buildSheetAmountValidation_v628_(salesAgg);
}

function buildSourceBasisValidation_v628_(salesAgg) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('원본대상행_검증') || SpreadsheetApp.getActive().insertSheet('원본대상행_검증');
  var rows = [
    ['항목','값','메모'],
    ['원본시트', salesAgg.sourceSheetName, '단일 원천'],
    ['AC열 헤더', salesAgg.acHeader, '매입금액 원천'],
    ['원본 전체 행수', salesAgg.rawRowCount, '헤더 제외'],
    ['분석대상 행수', salesAgg.detailRows.length, '부가세/대시보드 집계 기준'],
    ['취소/반품 제외 행수', salesAgg.cancelRowCount, '상태값 제외'],
    ['분석제외 행수', salesAgg.analysisExcludedRowCount, '취소 외 제외'],
    ['총매출액', salesAgg.totalGrossSales, '원본 전체'],
    ['취소/반품매출', salesAgg.totalCancelSales, '상태값 제외'],
    ['분석제외매출', salesAgg.totalAnalysisExcludedSales, '취소 외 제외'],
    ['순수매출액', salesAgg.totalNetSales, '분석대상'],
    ['AC 전체 합계', salesAgg.totalRawPurchase, '원본 전체'],
    ['AC 취소/반품 제외 합계', salesAgg.totalRawPurchase - salesAgg.totalCancelPurchase, '취소만 제외'],
    ['AC 분석대상 합계', salesAgg.totalPurchase, '부가세 시트 비교 기준'],
    ['실제정산금액', salesAgg.totalActualSettlement, '정산예정금액 입력'],
    ['예상정산금액(미정산)', salesAgg.totalEstimatedSettlement, '미정산 순수매출액 × 0.901'],
    ['정산기준금액', salesAgg.totalSettlementBasis, '실제 + 예상'],
    ['예상이익', salesAgg.totalEstimatedProfit, '정산기준금액 - AC 분석대상 합계']
  ];
  sheet.clearContents();
  sheet.getRange(1,1,rows.length,3).setValues(rows);
  formatOutputSheetBasic_v628_(sheet);
}

function buildExcludedRowsSheet_v628_(salesAgg) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('분석제외_행_리스트') || SpreadsheetApp.getActive().insertSheet('분석제외_행_리스트');
  var headers = ['원본행','제외사유','주문번호','상품번호','브랜드명','구매사이트명','마켓상태','더망고상태','매출금액','매입금액','상품명'];
  var rows = salesAgg.excludedRows.slice(0, 1000).map(function(r) { return [r.sheetRow,r.excludeReason,r.orderNo,r.productNo,r.brand,r.purchaseSite,r.marketStatus,r.themangoStatus,r.sales,r.purchase,compactText_v628_(r.productName,90)]; });
  replaceSheetValuesSafe_v628_(sheet, headers, rows);
  formatOutputSheetBasic_v628_(sheet);
}

function buildSheetAmountValidation_v628_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var headers = ['시트명','기준 순수매출액','시트 순수매출액','매출 차이','기준 매입금액','시트 매입금액','매입 차이','기준 정산기준금액','시트 정산기준금액','정산 차이','판정','메모'];
  var targets = [
    ['대시보드','summary'],
    ['브랜드별_마진율','table'],
    ['부가세_신고자료','table'],
    ['부가세_상품별','table'],
    ['부가세_고객별','table'],
    ['부가세_주문번호별','table']
  ];
  var rows = targets.map(function(t) {
    var actual = t[1] === 'summary' ? readDashboardSummaryAmounts_v628_(ss.getSheetByName(t[0])) : readTableAmounts_v628_(ss.getSheetByName(t[0]));
    var salesDiff = Math.round(salesAgg.totalNetSales - actual.sales);
    var purchaseDiff = Math.round(salesAgg.totalPurchase - actual.purchase);
    var settlementDiff = Math.round(salesAgg.totalSettlementBasis - actual.settlement);
    var ok = salesDiff === 0 && purchaseDiff === 0 && (t[0].indexOf('부가세_') === 0 ? true : settlementDiff === 0);
    return [t[0], salesAgg.totalNetSales, actual.sales, salesDiff, salesAgg.totalPurchase, actual.purchase, purchaseDiff, salesAgg.totalSettlementBasis, actual.settlement, settlementDiff, ok ? 'OK' : '차이확인', actual.memo || ''];
  });
  var sheet = ss.getSheetByName('시트별_금액검증') || ss.insertSheet('시트별_금액검증');
  replaceSheetValuesSafe_v628_(sheet, headers, rows);
  formatOutputSheetBasic_v628_(sheet);
}

function readDashboardSummaryAmounts_v628_(sheet) {
  var out = { sales:0, purchase:0, settlement:0, memo:'' };
  if (!sheet || sheet.getLastRow() < 1) { out.memo = '시트 없음'; return out; }
  var values = sheet.getDataRange().getValues();
  values.forEach(function(row) {
    var item = normText_v628_(row[1]);
    if (item === '순수매출액') out.sales = num_v628_(row[2]);
    if (item === '매입금액') out.purchase = num_v628_(row[2]);
    if (item === '정산기준금액') out.settlement = num_v628_(row[2]);
  });
  return out;
}

function readTableAmounts_v628_(sheet) {
  var out = { sales:0, purchase:0, settlement:0, memo:'' };
  if (!sheet || sheet.getLastRow() < 2) { out.memo = '시트 없음/데이터 없음'; return out; }
  var values = sheet.getDataRange().getValues();
  var header = values[0].map(normText_v628_);
  var salesIdx = findHeaderIndex_v628_(header, ['순수매출액']);
  var purchaseIdx = findHeaderIndex_v628_(header, ['매입금액']);
  var settlementIdx = findHeaderIndex_v628_(header, ['정산기준금액']);
  for (var r=1; r<values.length; r++) {
    if (salesIdx >= 0) out.sales += num_v628_(values[r][salesIdx]);
    if (purchaseIdx >= 0) out.purchase += num_v628_(values[r][purchaseIdx]);
    if (settlementIdx >= 0) out.settlement += num_v628_(values[r][settlementIdx]);
  }
  out.memo = 'headers sales=' + salesIdx + ' purchase=' + purchaseIdx + ' settlement=' + settlementIdx;
  return out;
}

function buildDashboardBrandRows_v628_(salesAgg, filterAgg) {
  var rows = [];
  Object.keys(salesAgg.byBrand).forEach(function(k) {
    var b = salesAgg.byBrand[k];
    var f = findFilterBrandInfo_v628_(filterAgg, k, b.brand);
    var collect = f.collectCount || 0;
    rows.push({
      brand:b.brand, filterName:f.filterName || '', accountId:f.accountId || '', collectCount:collect,
      salesProductCount:b.salesProductCount, orderCount:b.orderCount, netSalesAmount:b.netSalesAmount,
      actualSettlementAmount:b.actualSettlementAmount, estimatedSettlementAmount:b.estimatedSettlementAmount,
      settlementBasisAmount:b.settlementBasisAmount, purchaseAmount:b.purchaseAmount,
      estimatedProfitAmount:b.estimatedProfitAmount, estimatedProfitRate:b.estimatedProfitRate,
      unsettledOrderCount:b.unsettledOrderCount, overdueUnsettledOrderCount:b.overdueUnsettledOrderCount,
      latestRecentDate:f.latestRecentDate || '', unsoldCount:Math.max(0, collect - b.salesProductCount)
    });
  });
  return rows;
}

function findFilterBrandInfo_v628_(filterAgg, key, brand) {
  var empty = { collectCount:0, filterName:'', accountId:'', latestRecentDate:'' };
  if (!filterAgg) return empty;
  var byBrand = filterAgg.byBrand || filterAgg.brands || {};
  if (byBrand[key]) return byBrand[key];
  var altKey = brandKey_v628_(brand);
  if (byBrand[altKey]) return byBrand[altKey];
  return empty;
}

function splitVat_v628_(amount) { amount = Math.round(num_v628_(amount)); var supply = Math.round(amount / 1.1); return { total:amount, supply:supply, vat:amount-supply }; }
function vatPart_v628_(amount) { return splitVat_v628_(amount).vat; }
function rowToObject_v628_(headers, row) { var o={}; headers.forEach(function(h,i){ if(h) o[h]=row[i]; }); return o; }
function getAny_v628_(obj, names) { for (var i=0;i<names.length;i++) if (obj.hasOwnProperty(names[i]) && obj[names[i]] !== '' && obj[names[i]] != null) return obj[names[i]]; return ''; }
function num_v628_(v) { if (typeof toNumber_ === 'function') return toNumber_(v); if (typeof v === 'number') return v; var s=String(v==null?'':v).replace(/₩/g,'').replace(/,/g,'').trim(); var m=s.match(/-?\d+(\.\d+)?/); return m?Number(m[0]):0; }
function normalizeId_v628_(v) { return String(v==null?'':v).replace(/,/g,'').trim(); }
function normText_v628_(v) { return String(v==null?'':v).replace(/\n/g,'').replace(/\s+/g,'').trim(); }
function brandKey_v628_(brand) { if (typeof normalizeBrandKey_v613_ === 'function') return normalizeBrandKey_v613_(brand); return String(brand||'').replace(/\s+/g,'').toLowerCase(); }
function isCancelStatus_v628_(marketStatus, themangoStatus) { return /취소|반품|교환/.test(String(marketStatus||'') + ' ' + String(themangoStatus||'')); }
function parseDateValue_v628_(v) { if (Object.prototype.toString.call(v)==='[object Date]' && !isNaN(v)) return v; if (typeof v === 'number' && v > 20000) return new Date(Math.round((v - 25569) * 86400 * 1000)); var s=String(v||'').trim(); var m=s.match(/(20\d{2})[-.\/](\d{1,2})[-.\/](\d{1,2})/); if(m) return new Date(Number(m[1]),Number(m[2])-1,Number(m[3])); return null; }
function formatDateMMDD_v628_(d) { return d ? Utilities.formatDate(d, CONFIG.TIMEZONE, 'MM/dd') : ''; }
function formatNumberForMsg_v628_(n) { return String(Math.round(num_v628_(n))).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function compactText_v628_(text, limit) { text=String(text||'').replace(/\s+/g,' ').trim(); return text.length>limit ? text.slice(0,limit-1)+'…' : text; }
function findHeaderIndex_v628_(headers, names) { for (var i=0;i<headers.length;i++) for (var j=0;j<names.length;j++) if (headers[i]===normText_v628_(names[j])) return i; return -1; }
function decideNextRotationSafe_v628_(b) { if (typeof decideNextRotation_v611_ === 'function') return decideNextRotation_v611_(b); return b.unsoldCount > 0 ? '상품갈이 검토' : '유지'; }
function formatShortDateSafe_v628_(d) { if (!d) return ''; if (typeof formatShortDate_ === 'function') return formatShortDate_(d); return d; }

function replaceSheetValuesSafe_v628_(sheet, headers, rows) {
  if (typeof replaceSheetValues_v611_ === 'function') return replaceSheetValues_v611_(sheet, headers, rows);
  sheet.clearContents();
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  if (rows && rows.length) sheet.getRange(2,1,rows.length,headers.length).setValues(rows);
}

function formatOutputSheetBasic_v628_(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  try {
    sheet.getRange(1,1,1,lastCol).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    sheet.setFrozenRows(1);
    var headers = sheet.getRange(1,1,1,lastCol).getValues()[0].map(normText_v628_);
    for (var c=0;c<headers.length;c++) {
      var h = headers[c];
      var range = lastRow > 1 ? sheet.getRange(2,c+1,lastRow-1,1) : null;
      if (!range) continue;
      if (/율/.test(h)) range.setNumberFormat('0.00%');
      else if (/금액|매출|정산|매입|이익|수수료|부가세|공급/.test(h) && !/상품수|건수|수량/.test(h)) range.setNumberFormat('#,##0');
      else if (/수량|건수|상품수|브랜드수|행수/.test(h)) range.setNumberFormat('#,##0');
      else if (/날짜/.test(h)) range.setNumberFormat('@');
      else if (/주문번호|상품번호|계정ID/.test(h)) range.setNumberFormat('@');
    }
    if (typeof applyDisplayStandardsOnlyFast_v623 === 'function') {
      // do not call here to avoid recursive UI alerts; fast formatting is called by menu after generation
    }
  } catch (e) {}
}
