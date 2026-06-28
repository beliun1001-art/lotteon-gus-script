/**
 * LOTTEON v6.16 unsettled settlement / estimated profit patch
 *
 * 확정 기준:
 * - 매출데이터_붙여넣기 또는 매출데이터_정리의 정산예정금액(원)이 공란인 정상 주문은 미정산으로 봅니다.
 * - 미정산 정상 주문의 예상정산금액 = 순수매출액 × 0.901
 * - 실제정산금액 = 정산예정금액(원)이 있는 주문만 집계
 * - 정산기준금액 = 실제정산금액 + 예상정산금액
 * - 예상이익 = 정산기준금액 - 매입금액
 * - 미정산 주문은 비고에 '미정산' 표기
 * - 주문일로부터 30일 초과 미정산 주문은 대시보드 상단에 건수 표시하고, 쿠팡 계정별 별도 시트 생성
 */

var LOTTEON_PATCH_V616_UNSETTLED_ESTIMATED_PROFIT_LOADED = true;
var LOTTEON_ESTIMATED_SETTLEMENT_RATE_V616 = 0.901;
var LOTTEON_UNSETTLED_OVERDUE_DAYS_V616 = 30;

aggregateSalesByBrand_v611_ = function() {
  return aggregateSalesWithUnsettledEstimate_v616_();
};

rebuildDashboardBrandBased_v611_ = function(salesAgg, filterAgg) {
  return rebuildDashboardWithUnsettledEstimate_v616_(salesAgg || aggregateSalesWithUnsettledEstimate_v616_(), filterAgg || aggregateFiltersByBrand_v611_());
};

buildBrandMarginSheet_v611_ = function(salesAgg) {
  return buildBrandMarginSheetUnsettledEstimate_v616_(salesAgg || aggregateSalesWithUnsettledEstimate_v616_());
};

buildVatReportSheet_v611_ = function(salesAgg) {
  salesAgg = salesAgg || aggregateSalesWithUnsettledEstimate_v616_();
  if (typeof buildVatBreakdownSheetsNetSales_v613_ === 'function') buildVatBreakdownSheetsNetSales_v613_(salesAgg);
  buildUnsettledSettlementByAccountSheet_v616_(salesAgg);
  return null;
};

function aggregateSalesWithUnsettledEstimate_v616_() {
  var table = readSalesSourceRows_v613_();
  var rawPurchaseIndex = typeof buildRawPurchaseLookup_v614_ === 'function' ? buildRawPurchaseLookup_v614_() : null;
  var result = {
    byBrand: {},
    rows: table.rows,
    detailRows: [],
    sourceSheet: table.sourceSheet,
    hasPurchaseColumn: table.hasPurchaseColumn || (rawPurchaseIndex && rawPurchaseIndex.hasPurchaseColumn),
    totalGrossSales: 0,
    totalCancelSales: 0,
    totalNetSales: 0,
    totalActualSettlement: 0,
    totalEstimatedSettlement: 0,
    totalSettlementBasis: 0,
    totalPurchase: 0,
    totalMarketFee: 0,
    totalEstimatedProfit: 0,
    totalEstimatedProfitRate: 0,
    totalQty: 0,
    totalOrders: 0,
    totalProducts: 0,
    totalCustomers: 0,
    cancelQty: 0,
    cancelSales: 0,
    unsettledOrderSet: {},
    overdueUnsettledOrderSet: {},
    unsettledRows: [],
    overdueRows: [],
    totalUnsettledNetSales: 0,
    totalOverdueUnsettledNetSales: 0,
    // compatibility
    totalSales: 0,
    totalSettlement: 0,
    totalSettlementAmount: 0,
    totalNetProfit: 0,
    totalNetProfitRate: 0,
    totalMarketFeeRate: 0
  };

  var orderSet = {};
  var productSet = {};
  var customerSet = {};

  (table.rows || []).forEach(function(rowObj) {
    var brand = canonicalBrand_v613_(getFirstValue_v613_(rowObj, ['브랜드명_매칭','브랜드명','브랜드명_원본','브랜드','매출브랜드명']));
    if (!brand) brand = '브랜드미확인';
    var key = normalizeBrandKey_v613_(brand);

    var salesAmount = toNumber_(getFirstValue_v613_(rowObj, ['결제금액합계','결제금액','매출액','판매금액','총결제금액','상품금액']));
    var settlementRaw = getFirstValue_v616_(rowObj, ['정산예정금액(원)','정산예정금액','정산금액','정산액','실제정산금액']);
    var actualSettlement = toNumber_(settlementRaw);
    var hasActualSettlement = isNonBlankSettlement_v616_(settlementRaw) && actualSettlement !== 0;
    var qty = toNumber_(getFirstValue_v613_(rowObj, ['결제수량','수량','판매수량']));
    if (!qty && salesAmount) qty = 1;

    var orderNo = String(getFirstValue_v613_(rowObj, ['마켓주문번호','주문번호','주문ID','주문번호_원본']) || '').trim();
    var customer = String(getFirstValue_v613_(rowObj, ['고객명','수령인명','수령인','주문자명','구매자명','받는분','받는사람','receiverName']) || '').trim();
    var productNo = String(getFirstValue_v613_(rowObj, ['사이트상품번호','마켓상품번호','상품번호','상품ID','상품코드']) || '').trim();
    var productName = String(getFirstValue_v613_(rowObj, ['마켓상품명','원문상품명','상품명','상품명_원본','옵션명']) || '').trim();
    var accountId = String(getFirstValue_v613_(rowObj, ['쿠팡계정ID','원본계정ID','분석계정ID','계정ID','판매자ID']) || '').trim();
    var orderDateRaw = getFirstValue_v613_(rowObj, ['주문일시','주문일자','결제일시','결제일','주문일']);
    var dateParts = datePartsFromValue_v613_(orderDateRaw, getFirstValue_v613_(rowObj, ['주문월','월']));
    var orderDateObj = dateObjectFromDateParts_v616_(dateParts);
    var elapsedDays = orderDateObj ? Math.floor((startOfToday_v616_().getTime() - orderDateObj.getTime()) / 86400000) : '';
    var isOverdue = elapsedDays !== '' && elapsedDays > LOTTEON_UNSETTLED_OVERDUE_DAYS_V616;
    var cancel = isCancelRow_v613_(rowObj);

    var purchase = typeof getPurchaseAmountFromRow_v614_ === 'function' ? toNumber_(getPurchaseAmountFromRow_v614_(rowObj)) : toNumber_(getFirstValue_v613_(rowObj, ['매입금액','구매금액','매입가','구매가','원가','상품매입금액','매입금액합계']));
    if (!purchase && rawPurchaseIndex && typeof lookupPurchaseAmount_v614_ === 'function') {
      purchase = toNumber_(lookupPurchaseAmount_v614_(rawPurchaseIndex, orderNo, productNo, productName, salesAmount, qty));
    }

    if (!result.byBrand[key]) {
      result.byBrand[key] = createBrandAggRow_v616_(brand, key);
    }
    var b = result.byBrand[key];
    b.grossSalesAmount += salesAmount;
    result.totalGrossSales += salesAmount;

    if (cancel) {
      b.cancelSalesAmount += salesAmount;
      b.cancelQty += qty;
      result.totalCancelSales += salesAmount;
      result.cancelSales += salesAmount;
      result.cancelQty += qty;
      return;
    }

    var estimatedSettlement = hasActualSettlement ? 0 : Math.round(salesAmount * LOTTEON_ESTIMATED_SETTLEMENT_RATE_V616);
    var settlementBasis = hasActualSettlement ? actualSettlement : estimatedSettlement;
    var marketFee = salesAmount - settlementBasis;
    var estimatedProfit = settlementBasis - purchase;
    var unsettledStatus = hasActualSettlement ? '' : '미정산';

    b.netSalesAmount += salesAmount;
    b.actualSettlementAmount += hasActualSettlement ? actualSettlement : 0;
    b.estimatedSettlementAmount += estimatedSettlement;
    b.settlementBasisAmount += settlementBasis;
    b.purchaseAmount += purchase;
    b.marketFeeAmount += marketFee;
    b.estimatedProfitAmount += estimatedProfit;
    b.quantity += qty;
    if (orderNo) { b.orderSet[orderNo] = true; orderSet[orderNo] = true; }
    if (productNo || productName) { b.productSet[(productNo || '') + '|' + (productName || '')] = true; productSet[(productNo || '') + '|' + (productName || '')] = true; }
    if (customer) { b.customerSet[customer] = true; customerSet[customer] = true; }
    if (!hasActualSettlement) {
      b.unsettledOrderSet[orderNo || ('ROW_' + result.detailRows.length)] = true;
      result.unsettledOrderSet[orderNo || ('ROW_' + result.detailRows.length)] = true;
      b.unsettledNetSales += salesAmount;
      result.totalUnsettledNetSales += salesAmount;
      if (isOverdue) {
        b.overdueUnsettledOrderSet[orderNo || ('ROW_' + result.detailRows.length)] = true;
        result.overdueUnsettledOrderSet[orderNo || ('ROW_' + result.detailRows.length)] = true;
        result.totalOverdueUnsettledNetSales += salesAmount;
      }
    }

    var monthKey = dateParts.year && dateParts.month ? dateParts.year + '-' + ('0' + dateParts.month).slice(-2) : '';
    if (!b.months[monthKey]) b.months[monthKey] = createMonthAggRow_v616_(monthKey);
    var m = b.months[monthKey];
    m.sales += salesAmount;
    m.actualSettlement += hasActualSettlement ? actualSettlement : 0;
    m.estimatedSettlement += estimatedSettlement;
    m.settlementBasis += settlementBasis;
    m.purchase += purchase;
    m.marketFeeAmount += marketFee;
    m.estimatedProfitAmount += estimatedProfit;
    m.qty += qty;
    if (orderNo) m.orderSet[orderNo] = true;

    var detail = {
      year: dateParts.year,
      month: dateParts.month,
      day: dateParts.day,
      orderDate: dateParts.dateText,
      elapsedDays: elapsedDays,
      accountId: accountId,
      orderNo: orderNo,
      customer: customer,
      brand: brand,
      productNo: productNo,
      productName: productName,
      qty: qty,
      netSales: salesAmount,
      actualSettlement: hasActualSettlement ? actualSettlement : 0,
      estimatedSettlement: estimatedSettlement,
      settlementBasis: settlementBasis,
      settlement: settlementBasis,
      marketFee: marketFee,
      purchase: purchase,
      estimatedProfit: estimatedProfit,
      netProfit: estimatedProfit,
      unsettledStatus: unsettledStatus,
      overdueStatus: (!hasActualSettlement && isOverdue) ? '30일초과 미정산' : '',
      note: !hasActualSettlement ? '미정산' : '',
      sourceRow: rowObj
    };
    result.detailRows.push(detail);
    if (!hasActualSettlement) {
      result.unsettledRows.push(detail);
      if (isOverdue) result.overdueRows.push(detail);
    }

    result.totalNetSales += salesAmount;
    result.totalActualSettlement += hasActualSettlement ? actualSettlement : 0;
    result.totalEstimatedSettlement += estimatedSettlement;
    result.totalSettlementBasis += settlementBasis;
    result.totalPurchase += purchase;
    result.totalMarketFee += marketFee;
    result.totalEstimatedProfit += estimatedProfit;
    result.totalQty += qty;
  });

  Object.keys(result.byBrand).forEach(function(key) {
    finalizeBrandAggRow_v616_(result.byBrand[key]);
  });

  result.totalOrders = Object.keys(orderSet).length;
  result.totalProducts = Object.keys(productSet).length;
  result.totalCustomers = Object.keys(customerSet).length;
  result.totalEstimatedProfitRate = result.totalNetSales ? result.totalEstimatedProfit / result.totalNetSales : 0;
  result.totalMarketFeeRate = result.totalNetSales ? result.totalMarketFee / result.totalNetSales : 0;

  // compatibility
  result.totalSales = result.totalNetSales;
  result.totalSettlement = result.totalSettlementBasis;
  result.totalSettlementAmount = result.totalSettlementBasis;
  result.totalNetProfit = result.totalEstimatedProfit;
  result.totalNetProfitRate = result.totalEstimatedProfitRate;

  return result;
}

function createBrandAggRow_v616_(brand, key) {
  return {
    brand: brand,
    key: key,
    grossSalesAmount: 0,
    cancelSalesAmount: 0,
    netSalesAmount: 0,
    actualSettlementAmount: 0,
    estimatedSettlementAmount: 0,
    settlementBasisAmount: 0,
    purchaseAmount: 0,
    marketFeeAmount: 0,
    estimatedProfitAmount: 0,
    quantity: 0,
    cancelQty: 0,
    unsettledNetSales: 0,
    orderSet: {},
    productSet: {},
    customerSet: {},
    unsettledOrderSet: {},
    overdueUnsettledOrderSet: {},
    months: {},
    // compatibility
    salesAmount: 0,
    settlementAmount: 0,
    netProfitAmount: 0,
    marginAmount: 0,
    marginRate: 0,
    netSalesForVat: 0
  };
}

function createMonthAggRow_v616_(monthKey) {
  return { month: monthKey, sales: 0, actualSettlement: 0, estimatedSettlement: 0, settlementBasis: 0, purchase: 0, marketFeeAmount: 0, estimatedProfitAmount: 0, qty: 0, orderSet: {}, netSalesForVat: 0 };
}

function finalizeBrandAggRow_v616_(b) {
  b.orderCount = Object.keys(b.orderSet).length || (b.netSalesAmount ? 1 : 0);
  b.salesProductCount = Object.keys(b.productSet).length;
  b.customerCount = Object.keys(b.customerSet).length;
  b.unsettledOrderCount = Object.keys(b.unsettledOrderSet).length;
  b.overdueUnsettledOrderCount = Object.keys(b.overdueUnsettledOrderSet).length;
  b.marketFeeRate = b.netSalesAmount ? b.marketFeeAmount / b.netSalesAmount : 0;
  b.estimatedProfitRate = b.netSalesAmount ? b.estimatedProfitAmount / b.netSalesAmount : 0;

  // compatibility for prior dashboard helpers
  b.salesAmount = b.netSalesAmount;
  b.settlementAmount = b.settlementBasisAmount;
  b.netProfitAmount = b.estimatedProfitAmount;
  b.marginAmount = b.estimatedProfitAmount;
  b.marginRate = b.estimatedProfitRate;
  b.netSalesForVat = b.netSalesAmount;
}

function rebuildDashboardWithUnsettledEstimate_v616_(salesAgg, filterAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD) || ss.insertSheet(CONFIG.SHEETS.DASHBOARD);
  salesAgg = salesAgg || aggregateSalesWithUnsettledEstimate_v616_();
  filterAgg = filterAgg || aggregateFiltersByBrand_v611_();
  var brandRows = buildBrandRowsForDashboard_v611_(salesAgg, filterAgg);
  brandRows.forEach(function(b) {
    var source = salesAgg.byBrand[b.key] || {};
    b.grossSalesAmount = source.grossSalesAmount || 0;
    b.cancelSalesAmount = source.cancelSalesAmount || 0;
    b.netSalesAmount = source.netSalesAmount || b.salesAmount || 0;
    b.actualSettlementAmount = source.actualSettlementAmount || 0;
    b.estimatedSettlementAmount = source.estimatedSettlementAmount || 0;
    b.settlementBasisAmount = source.settlementBasisAmount || b.settlementAmount || 0;
    b.marketFeeAmount = source.marketFeeAmount || 0;
    b.estimatedProfitAmount = source.estimatedProfitAmount || 0;
    b.estimatedProfitRate = source.estimatedProfitRate || 0;
    b.unsettledOrderCount = source.unsettledOrderCount || 0;
    b.overdueUnsettledOrderCount = source.overdueUnsettledOrderCount || 0;
    b.unsettledNetSales = source.unsettledNetSales || 0;
    b.salesAmount = b.netSalesAmount;
  });

  var nowText = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  var overdueCount = Object.keys(salesAgg.overdueUnsettledOrderSet || {}).length;
  var unsettledCount = Object.keys(salesAgg.unsettledOrderSet || {}).length;
  var rows = [];
  rows.push(['요약','갱신기준','브랜드 기준 + 미정산 예상정산 90.1% 반영','','','','','','','','v6.16']);
  rows.push(['요약','갱신시각',nowText,'','','','','','','','Asia/Seoul']);
  rows.push(['요약','분석브랜드수',brandRows.length,'','','','','','','','브랜드명 기준']);
  rows.push(['요약','수집상품수',filterAgg.totalCollectCount,'','','','','','','','필터별_상품수 API_totalCount 총합']);
  rows.push(['요약','매출상품수',salesAgg.totalProducts,'','','','','','','','취소/반품 제외']);
  rows.push(['요약','주문건수',salesAgg.totalOrders,'','','','','','','','취소/반품 제외']);
  rows.push(['요약','미정산 주문건수',unsettledCount,'','','','','','','','정산예정금액 공란 정상 주문']);
  rows.push(['요약','30일초과 미정산 주문건수',overdueCount,'','','','','','','','주문일 기준 30일 초과']);
  rows.push(['요약','총매출액',salesAgg.totalGrossSales,'','','','','','','','취소/반품 포함']);
  rows.push(['요약','취소/반품매출',salesAgg.totalCancelSales,'','','','','','','','취소/반품 판정 row']);
  rows.push(['요약','순수매출액',salesAgg.totalNetSales,'','','','','','','','취소/반품 제외']);
  rows.push(['요약','실제정산금액',salesAgg.totalActualSettlement,'','','','','','','','정산예정금액 입력 완료 주문']);
  rows.push(['요약','예상정산금액(미정산)',salesAgg.totalEstimatedSettlement,'','','','','','','','미정산 순수매출액 × 0.901']);
  rows.push(['요약','정산기준금액',salesAgg.totalSettlementBasis,'','','','','','','','실제정산금액 + 예상정산금액']);
  rows.push(['요약','마켓수수료/비용',salesAgg.totalMarketFee,'','','','','','','','순수매출액 - 정산기준금액']);
  rows.push(['요약','마켓수수료율',salesAgg.totalMarketFeeRate,'','','','','','','','마켓수수료 ÷ 순수매출액']);
  rows.push(['요약','매입금액',salesAgg.totalPurchase,'','','','','','','',salesAgg.hasPurchaseColumn ? '취소제외 매입/원가 컬럼 기준' : '매입/원가 컬럼 미확인']);
  rows.push(['요약','예상이익',salesAgg.totalEstimatedProfit,'','','','','','','','정산기준금액 - 매입금액']);
  rows.push(['요약','예상이익률',salesAgg.totalEstimatedProfitRate,'','','','','','','','예상이익 ÷ 순수매출액']);
  var vatBaseTotal = Math.round(salesAgg.totalNetSales / 1.1);
  rows.push(['요약','부가세 신고 공급대가',salesAgg.totalNetSales,'','','','','','','','순수매출액 기준']);
  rows.push(['요약','부가세 신고 공급가액',vatBaseTotal,'','','','','','','','공급대가 ÷ 1.1']);
  rows.push(['요약','부가세',salesAgg.totalNetSales - vatBaseTotal,'','','','','','','','공급대가 - 공급가액']);
  rows.push(['','','','','','','','','','','']);

  rows.push(['구분','브랜드명','대표검색필터명','계정ID','수집수','매출상품','주문','순수매출액','정산기준금액','예상이익률','메모']);
  brandRows.slice().sort(function(a, b) { return (b.netSalesAmount || 0) - (a.netSalesAmount || 0); }).slice(0, 30).forEach(function(b) {
    rows.push(['브랜드TOP', b.brand, b.filterName, b.accountId, b.collectCount, b.salesProductCount, b.orderCount, b.netSalesAmount, b.settlementBasisAmount, b.estimatedProfitRate, '실제정산 ' + formatWonText_v613_(b.actualSettlementAmount) + ' / 미정산예상 ' + formatWonText_v613_(b.estimatedSettlementAmount) + ' / 미정산 ' + b.unsettledOrderCount + '건 / 예상이익 ' + formatWonText_v613_(b.estimatedProfitAmount)]);
  });
  rows.push(['','','','','','','','','','','']);

  rows.push(['구분','브랜드명','대표검색필터명','계정ID','수집수','매출상품','주문','순수매출액','정산기준금액','다음작업','상품갈이메모']);
  brandRows.slice().sort(function(a, b) { return (b.unsoldCount || 0) - (a.unsoldCount || 0); }).slice(0, 40).forEach(function(b) {
    rows.push(['상품갈이', b.brand, b.filterName, b.accountId, b.collectCount, b.salesProductCount, b.orderCount, b.netSalesAmount, b.settlementBasisAmount, decideNextRotation_v611_(b), '최근수집 ' + formatShortDate_(b.latestRecentDate) + ' / 미판매 ' + b.unsoldCount + ' / 미정산 ' + b.unsettledOrderCount + '건']);
  });

  replaceSheetValues_v611_(sheet, CONFIG.HEADERS.DASHBOARD, rows);
  formatDashboardSettlementProfit_v612_(sheet);
  if (typeof applyDashboardDynamicHeaderStyle_v615_ === 'function') applyDashboardDynamicHeaderStyle_v615_();
  buildUnsettledSettlementByAccountSheet_v616_(salesAgg);
}

function buildBrandMarginSheetUnsettledEstimate_v616_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('브랜드별_마진율') || ss.insertSheet('브랜드별_마진율');
  var headers = ['브랜드명','주문건수','고객수','판매수량','매출상품수','총매출액','취소/반품매출','순수매출액','실제정산금액','예상정산금액(미정산)','정산기준금액','마켓수수료/비용','마켓수수료율','매입금액','예상이익','예상이익률','미정산주문건수','30일초과미정산건수','비고'];
  var rows = Object.keys(salesAgg.byBrand || {}).map(function(key) {
    var b = salesAgg.byBrand[key];
    return [
      b.brand,
      b.orderCount,
      b.customerCount,
      b.quantity,
      b.salesProductCount,
      b.grossSalesAmount,
      b.cancelSalesAmount,
      b.netSalesAmount,
      b.actualSettlementAmount,
      b.estimatedSettlementAmount,
      b.settlementBasisAmount,
      b.marketFeeAmount,
      b.marketFeeRate,
      b.purchaseAmount,
      b.estimatedProfitAmount,
      b.estimatedProfitRate,
      b.unsettledOrderCount,
      b.overdueUnsettledOrderCount,
      b.unsettledOrderCount ? '미정산 포함: 미정산 주문은 순수매출액×0.901 예상정산 적용' : '정산완료 기준'
    ];
  }).sort(function(a, b) { return b[7] - a[7]; });
  replaceSheetValues_v611_(sheet, headers, rows);
  formatNumberColumns_v611_(sheet, [2,3,4,5,17,18]);
  formatMoneyColumns_v611_(sheet, [6,7,8,9,10,11,12,14,15]);
  if (rows.length) {
    sheet.getRange(2, 13, rows.length, 1).setNumberFormat('0.00%');
    sheet.getRange(2, 16, rows.length, 1).setNumberFormat('0.00%');
  }
}

function buildUnsettledSettlementByAccountSheet_v616_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('미정산_쿠팡계정별') || ss.insertSheet('미정산_쿠팡계정별');
  var headers = ['쿠팡계정ID','구분','미정산건수','30일초과건수','주문일','경과일','주문번호','브랜드명','고객명','상품번호','상품명','순수매출액','예상정산금액','매입금액','예상이익','비고'];
  var summary = {};
  (salesAgg.unsettledRows || []).forEach(function(d) {
    var account = d.accountId || '계정미확인';
    if (!summary[account]) summary[account] = { count: 0, overdue: 0, netSales: 0, estimated: 0, purchase: 0, profit: 0, orders: {} };
    var s = summary[account];
    s.count++;
    if (d.overdueStatus) s.overdue++;
    s.netSales += d.netSales || 0;
    s.estimated += d.estimatedSettlement || 0;
    s.purchase += d.purchase || 0;
    s.profit += d.estimatedProfit || 0;
    if (d.orderNo) s.orders[d.orderNo] = true;
  });

  var rows = [];
  Object.keys(summary).sort().forEach(function(account) {
    var s = summary[account];
    rows.push([account, '요약', s.count, s.overdue, '', '', Object.keys(s.orders).join(', '), '', '', '', '', s.netSales, s.estimated, s.purchase, s.profit, '미정산 주문 요약']);
  });
  if (rows.length) rows.push(['','','','','','','','','','','','','','','','']);
  (salesAgg.unsettledRows || []).slice().sort(function(a, b) {
    return String(a.accountId || '').localeCompare(String(b.accountId || '')) || String(a.orderDate || '').localeCompare(String(b.orderDate || '')) || String(a.orderNo || '').localeCompare(String(b.orderNo || ''));
  }).forEach(function(d) {
    rows.push([d.accountId || '계정미확인', '상세', '', d.overdueStatus ? 1 : 0, d.orderDate, d.elapsedDays, d.orderNo, d.brand, d.customer, d.productNo, d.productName, d.netSales, d.estimatedSettlement, d.purchase, d.estimatedProfit, d.overdueStatus || '미정산']);
  });

  replaceSheetValues_v611_(sheet, headers, rows);
  formatNumberColumns_v611_(sheet, [3,4,6]);
  formatMoneyColumns_v611_(sheet, [12,13,14,15]);
  try {
    sheet.getRange(2, 7, Math.max(sheet.getLastRow() - 1, 1), 1).setNumberFormat('@');
  } catch (e) {}
}

function getFirstValue_v616_(obj, names) {
  for (var i = 0; i < names.length; i++) {
    if (Object.prototype.hasOwnProperty.call(obj, names[i])) return obj[names[i]];
  }
  return '';
}

function isNonBlankSettlement_v616_(v) {
  if (v == null) return false;
  if (Object.prototype.toString.call(v) === '[object Date]') return false;
  var text = String(v).replace(/[₩,\s]/g, '').trim();
  if (!text) return false;
  if (/^-+$/.test(text)) return false;
  return !isNaN(Number(text));
}

function dateObjectFromDateParts_v616_(parts) {
  if (!parts || !parts.year || !parts.month || !parts.day) return null;
  var d = new Date(Number(parts.year), Number(parts.month) - 1, Number(parts.day));
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function startOfToday_v616_() {
  var now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}
