/**
 * LOTTEON v6.13 net-sales profit and VAT breakdown patch
 *
 * 확정 기준:
 * - 총매출액 = 취소/반품 포함 원본 매출 합계
 * - 순수매출액 = 취소/반품 제외 매출 합계
 * - 실제정산금액 = 정산예정금액 중 취소/반품 제외 합계
 * - 마켓수수료/비용 = 순수매출액 - 실제정산금액
 * - 순수익 = 실제정산금액 - 매입금액
 * - 순수익률 = 순수익 / 순수매출액
 * - 부가세 신고자료는 총매출이 아니라 순수매출액 기준으로 계산
 *
 * 출력:
 * - 브랜드별_마진율: 총매출액/취소반품매출/순수매출액 분리
 * - 부가세_신고자료: 년/월/일 + 주문번호/고객/상품 상세 기준
 * - 부가세_상품별, 부가세_고객별, 부가세_주문번호별: 순수매출 기준 요약
 */

var LOTTEON_PATCH_V613_NET_SALES_VAT_BREAKDOWN_LOADED = true;

var __baseRebuildDashboardBrandBased_v613 = typeof rebuildDashboardBrandBased_v611_ === 'function' ? rebuildDashboardBrandBased_v611_ : null;
var __baseBuildBrandMarginSheet_v613 = typeof buildBrandMarginSheet_v611_ === 'function' ? buildBrandMarginSheet_v611_ : null;
var __baseBuildVatReportSheet_v613 = typeof buildVatReportSheet_v611_ === 'function' ? buildVatReportSheet_v611_ : null;

aggregateSalesByBrand_v611_ = function() {
  return aggregateSalesNetByBrand_v613_();
};

rebuildDashboardBrandBased_v611_ = function(salesAgg, filterAgg) {
  return rebuildDashboardBrandBasedNetSales_v613_(salesAgg || aggregateSalesNetByBrand_v613_(), filterAgg || aggregateFiltersByBrand_v611_());
};

buildBrandMarginSheet_v611_ = function(salesAgg) {
  return buildBrandMarginSheetNetSales_v613_(salesAgg || aggregateSalesNetByBrand_v613_());
};

buildVatReportSheet_v611_ = function(salesAgg) {
  return buildVatBreakdownSheetsNetSales_v613_(salesAgg || aggregateSalesNetByBrand_v613_());
};

function aggregateSalesNetByBrand_v613_() {
  var table = readSalesSourceRows_v613_();
  var result = {
    byBrand: {},
    rows: table.rows,
    detailRows: [],
    sourceSheet: table.sourceSheet,
    hasPurchaseColumn: table.hasPurchaseColumn,
    totalGrossSales: 0,
    totalCancelSales: 0,
    totalNetSales: 0,
    totalSettlement: 0,
    totalPurchase: 0,
    totalMarketFee: 0,
    totalNetProfit: 0,
    totalNetProfitRate: 0,
    totalMarketFeeRate: 0,
    totalQty: 0,
    totalOrders: 0,
    totalProducts: 0,
    cancelQty: 0,
    cancelSales: 0,
    // v6.11/v6.12 compatibility
    totalSales: 0,
    totalSettlementAmount: 0
  };

  var orderSet = {};
  var productSet = {};

  (table.rows || []).forEach(function(rowObj) {
    var brand = canonicalBrand_v613_(getFirstValue_v613_(rowObj, ['브랜드명_매칭','브랜드명','브랜드명_원본','브랜드','매출브랜드명']));
    if (!brand) brand = '브랜드미확인';
    var key = normalizeBrandKey_v613_(brand);
    var salesAmount = toNumber_(getFirstValue_v613_(rowObj, ['결제금액합계','결제금액','매출액','판매금액','총결제금액','상품금액']));
    var settlement = toNumber_(getFirstValue_v613_(rowObj, ['정산예정금액','정산금액','정산액']));
    var purchase = toNumber_(getFirstValue_v613_(rowObj, ['매입금액','구매금액','매입가','구매가','원가','상품매입금액','매입금액합계']));
    var qty = toNumber_(getFirstValue_v613_(rowObj, ['결제수량','수량','판매수량']));
    if (!qty && salesAmount) qty = 1;

    var orderNo = String(getFirstValue_v613_(rowObj, ['마켓주문번호','주문번호','주문ID','주문번호_원본']) || '').trim();
    var customer = String(getFirstValue_v613_(rowObj, ['고객명','수령인명','수령인','주문자명','구매자명','받는분','받는사람','receiverName']) || '').trim();
    var productNo = String(getFirstValue_v613_(rowObj, ['사이트상품번호','마켓상품번호','상품번호','상품ID','상품코드']) || '').trim();
    var productName = String(getFirstValue_v613_(rowObj, ['마켓상품명','원문상품명','상품명','상품명_원본','옵션명']) || '').trim();
    var orderDateRaw = getFirstValue_v613_(rowObj, ['주문일시','주문일자','결제일시','결제일','주문일']);
    var dateParts = datePartsFromValue_v613_(orderDateRaw, getFirstValue_v613_(rowObj, ['주문월','월']));
    var cancel = isCancelRow_v613_(rowObj);

    if (!result.byBrand[key]) {
      result.byBrand[key] = {
        brand: brand,
        key: key,
        grossSalesAmount: 0,
        cancelSalesAmount: 0,
        netSalesAmount: 0,
        settlementAmount: 0,
        purchaseAmount: 0,
        marketFeeAmount: 0,
        marketFeeRate: 0,
        netProfitAmount: 0,
        netProfitRate: 0,
        quantity: 0,
        cancelQty: 0,
        orderSet: {},
        productSet: {},
        customerSet: {},
        orderCount: 0,
        salesProductCount: 0,
        customerCount: 0,
        months: {},
        // compatibility
        salesAmount: 0,
        marginAmount: 0,
        marginRate: 0,
        netSalesForVat: 0
      };
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

    b.netSalesAmount += salesAmount;
    b.settlementAmount += settlement;
    b.purchaseAmount += purchase;
    b.quantity += qty;
    b.netSalesForVat += salesAmount;
    if (orderNo) { b.orderSet[orderNo] = true; orderSet[orderNo] = true; }
    if (productNo || productName) { b.productSet[(productNo || '') + '|' + (productName || '')] = true; productSet[(productNo || '') + '|' + (productName || '')] = true; }
    if (customer) b.customerSet[customer] = true;

    var monthKey = dateParts.year && dateParts.month ? dateParts.year + '-' + ('0' + dateParts.month).slice(-2) : '';
    if (!b.months[monthKey]) b.months[monthKey] = { month: monthKey, sales: 0, settlement: 0, purchase: 0, qty: 0, orderSet: {}, netSalesForVat: 0, marketFeeAmount: 0, netProfitAmount: 0 };
    var m = b.months[monthKey];
    m.sales += salesAmount;
    m.settlement += settlement;
    m.purchase += purchase;
    m.qty += qty;
    m.netSalesForVat += salesAmount;
    if (orderNo) m.orderSet[orderNo] = true;

    result.totalNetSales += salesAmount;
    result.totalSettlement += settlement;
    result.totalPurchase += purchase;
    result.totalQty += qty;

    result.detailRows.push({
      year: dateParts.year,
      month: dateParts.month,
      day: dateParts.day,
      orderDate: dateParts.dateText,
      orderNo: orderNo,
      customer: customer,
      brand: brand,
      productNo: productNo,
      productName: productName,
      qty: qty,
      netSales: salesAmount,
      settlement: settlement,
      purchase: purchase,
      marketFee: salesAmount - settlement,
      netProfit: settlement - purchase,
      sourceRow: rowObj
    });
  });

  Object.keys(result.byBrand).forEach(function(key) {
    var b = result.byBrand[key];
    b.orderCount = Object.keys(b.orderSet).length || (b.netSalesAmount ? 1 : 0);
    b.salesProductCount = Object.keys(b.productSet).length;
    b.customerCount = Object.keys(b.customerSet).length;
    b.marketFeeAmount = b.netSalesAmount - b.settlementAmount;
    b.marketFeeRate = b.netSalesAmount ? b.marketFeeAmount / b.netSalesAmount : 0;
    b.netProfitAmount = b.settlementAmount - b.purchaseAmount;
    b.netProfitRate = b.netSalesAmount ? b.netProfitAmount / b.netSalesAmount : 0;

    // compatibility for old dashboard helpers
    b.salesAmount = b.netSalesAmount;
    b.marginAmount = b.netProfitAmount;
    b.marginRate = b.netProfitRate;

    Object.keys(b.months || {}).forEach(function(month) {
      var m = b.months[month];
      m.marketFeeAmount = m.sales - m.settlement;
      m.netProfitAmount = m.settlement - m.purchase;
      m.marketFeeRate = m.sales ? m.marketFeeAmount / m.sales : 0;
      m.netProfitRate = m.sales ? m.netProfitAmount / m.sales : 0;
    });
  });

  result.totalOrders = Object.keys(orderSet).length;
  result.totalProducts = Object.keys(productSet).length;
  result.totalMarketFee = result.totalNetSales - result.totalSettlement;
  result.totalNetProfit = result.totalSettlement - result.totalPurchase;
  result.totalMarketFeeRate = result.totalNetSales ? result.totalMarketFee / result.totalNetSales : 0;
  result.totalNetProfitRate = result.totalNetSales ? result.totalNetProfit / result.totalNetSales : 0;

  // compatibility
  result.totalSales = result.totalNetSales;
  result.totalSettlementAmount = result.totalSettlement;

  return result;
}

function buildBrandMarginSheetNetSales_v613_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('브랜드별_마진율') || ss.insertSheet('브랜드별_마진율');
  var headers = ['브랜드명','주문건수','고객수','판매수량','매출상품수','총매출액','취소/반품매출','순수매출액','실제정산금액','마켓수수료/비용','마켓수수료율','매입금액','순수익','순수익률','비고'];
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
      b.settlementAmount,
      b.marketFeeAmount,
      b.marketFeeRate,
      b.purchaseAmount,
      b.netProfitAmount,
      b.netProfitRate,
      salesAgg.hasPurchaseColumn ? '순수익=실제정산금액-매입금액 / 순수익률=순수익÷순수매출액' : '매입/원가 컬럼 미확인: 순수익 검증 필요'
    ];
  }).sort(function(a, b) { return b[7] - a[7]; });

  replaceSheetValues_v611_(sheet, headers, rows);
  formatNumberColumns_v611_(sheet, [2,3,4,5]);
  formatMoneyColumns_v611_(sheet, [6,7,8,9,10,12,13]);
  if (rows.length) {
    sheet.getRange(2, 11, rows.length, 1).setNumberFormat('0.00%');
    sheet.getRange(2, 14, rows.length, 1).setNumberFormat('0.00%');
  }
}

function buildVatBreakdownSheetsNetSales_v613_(salesAgg) {
  buildVatDetailSheet_v613_(salesAgg);
  buildVatProductSheet_v613_(salesAgg);
  buildVatCustomerSheet_v613_(salesAgg);
  buildVatOrderSheet_v613_(salesAgg);
}

function buildVatDetailSheet_v613_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('부가세_신고자료') || ss.insertSheet('부가세_신고자료');
  var headers = ['년','월','일','주문번호','고객명','브랜드명','상품번호','상품명','판매수량','순수매출액','공급가액','부가세','실제정산금액','마켓수수료/비용','매입금액','순수익','비고'];
  var rows = (salesAgg.detailRows || []).map(function(d) {
    var base = Math.round((d.netSales || 0) / 1.1);
    var vat = (d.netSales || 0) - base;
    return [d.year, d.month, d.day, d.orderNo, d.customer, d.brand, d.productNo, d.productName, d.qty, d.netSales, base, vat, d.settlement, d.marketFee, d.purchase, d.netProfit, '취소/반품 제외 순수매출 기준'];
  }).sort(sortVatRows_v613_);
  replaceSheetValues_v611_(sheet, headers, rows);
  formatNumberColumns_v611_(sheet, [1,2,3,9]);
  formatMoneyColumns_v611_(sheet, [10,11,12,13,14,15,16]);
}

function buildVatProductSheet_v613_(salesAgg) {
  var grouped = groupVatDetails_v613_(salesAgg.detailRows || [], function(d) { return [d.brand, d.productNo, d.productName].join('|'); }, function(d) { return { brand: d.brand, productNo: d.productNo, productName: d.productName, orderSet: {}, customerSet: {} }; });
  Object.keys(grouped).forEach(function(k) {
    var g = grouped[k];
    (g.rows || []).forEach(function(d) { if (d.orderNo) g.orderSet[d.orderNo] = true; if (d.customer) g.customerSet[d.customer] = true; });
  });
  var rows = Object.keys(grouped).map(function(k) {
    var g = grouped[k];
    var base = Math.round(g.netSales / 1.1);
    return [g.brand, g.productNo, g.productName, g.qty, g.netSales, base, g.netSales - base, g.settlement, g.marketFee, g.purchase, g.netProfit, Object.keys(g.orderSet).length, Object.keys(g.customerSet).length];
  }).sort(function(a, b) { return b[4] - a[4]; });
  writeVatSummarySheet_v613_('부가세_상품별', ['브랜드명','상품번호','상품명','판매수량','순수매출액','공급가액','부가세','실제정산금액','마켓수수료/비용','매입금액','순수익','주문건수','고객수'], rows, [5,6,7,8,9,10,11], [4,12,13]);
}

function buildVatCustomerSheet_v613_(salesAgg) {
  var grouped = groupVatDetails_v613_(salesAgg.detailRows || [], function(d) { return [d.customer || '고객명미확인'].join('|'); }, function(d) { return { customer: d.customer || '고객명미확인', orderSet: {}, productSet: {} }; });
  Object.keys(grouped).forEach(function(k) {
    var g = grouped[k];
    (g.rows || []).forEach(function(d) { if (d.orderNo) g.orderSet[d.orderNo] = true; if (d.productNo || d.productName) g.productSet[(d.productNo || '') + '|' + (d.productName || '')] = true; });
  });
  var rows = Object.keys(grouped).map(function(k) {
    var g = grouped[k];
    var base = Math.round(g.netSales / 1.1);
    return [g.customer, Object.keys(g.orderSet).length, Object.keys(g.productSet).length, g.qty, g.netSales, base, g.netSales - base, g.settlement, g.marketFee, g.purchase, g.netProfit];
  }).sort(function(a, b) { return b[4] - a[4]; });
  writeVatSummarySheet_v613_('부가세_고객별', ['고객명','주문건수','상품수','판매수량','순수매출액','공급가액','부가세','실제정산금액','마켓수수료/비용','매입금액','순수익'], rows, [5,6,7,8,9,10,11], [2,3,4]);
}

function buildVatOrderSheet_v613_(salesAgg) {
  var grouped = groupVatDetails_v613_(salesAgg.detailRows || [], function(d) { return d.orderNo || '주문번호미확인'; }, function(d) { return { year: d.year, month: d.month, day: d.day, orderNo: d.orderNo || '주문번호미확인', customer: d.customer, brandSet: {}, productSet: {} }; });
  Object.keys(grouped).forEach(function(k) {
    var g = grouped[k];
    (g.rows || []).forEach(function(d) { if (d.brand) g.brandSet[d.brand] = true; if (d.productNo || d.productName) g.productSet[(d.productNo || '') + '|' + (d.productName || '')] = true; if (!g.customer && d.customer) g.customer = d.customer; });
  });
  var rows = Object.keys(grouped).map(function(k) {
    var g = grouped[k];
    var base = Math.round(g.netSales / 1.1);
    return [g.year, g.month, g.day, g.orderNo, g.customer, Object.keys(g.brandSet).length, Object.keys(g.productSet).length, g.qty, g.netSales, base, g.netSales - base, g.settlement, g.marketFee, g.purchase, g.netProfit];
  }).sort(sortVatRows_v613_);
  writeVatSummarySheet_v613_('부가세_주문번호별', ['년','월','일','주문번호','고객명','브랜드수','상품수','판매수량','순수매출액','공급가액','부가세','실제정산금액','마켓수수료/비용','매입금액','순수익'], rows, [9,10,11,12,13,14,15], [1,2,3,6,7,8]);
}

function rebuildDashboardBrandBasedNetSales_v613_(salesAgg, filterAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD) || ss.insertSheet(CONFIG.SHEETS.DASHBOARD);
  salesAgg = salesAgg || aggregateSalesNetByBrand_v613_();
  filterAgg = filterAgg || aggregateFiltersByBrand_v611_();
  var brandRows = buildBrandRowsForDashboard_v611_(salesAgg, filterAgg);
  brandRows.forEach(function(b) {
    var source = salesAgg.byBrand[b.key] || {};
    b.grossSalesAmount = source.grossSalesAmount || 0;
    b.cancelSalesAmount = source.cancelSalesAmount || 0;
    b.netSalesAmount = source.netSalesAmount || b.salesAmount || 0;
    b.marketFeeAmount = b.netSalesAmount - (b.settlementAmount || 0);
    b.netProfitAmount = (b.settlementAmount || 0) - (b.purchaseAmount || 0);
    b.netProfitRate = b.netSalesAmount ? b.netProfitAmount / b.netSalesAmount : 0;
    b.salesAmount = b.netSalesAmount;
  });

  var nowText = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  var rows = [];
  rows.push(['요약','갱신기준','브랜드 기준 + 취소/반품 제외 순수매출 기준','','','','','','','','v6.13']);
  rows.push(['요약','갱신시각',nowText,'','','','','','','','Asia/Seoul']);
  rows.push(['요약','분석브랜드수',brandRows.length,'','','','','','','','브랜드명 기준']);
  rows.push(['요약','수집상품수',filterAgg.totalCollectCount,'','','','','','','','필터별_상품수 API_totalCount 총합']);
  rows.push(['요약','매출상품수',salesAgg.totalProducts,'','','','','','','','취소/반품 제외']);
  rows.push(['요약','주문건수',salesAgg.totalOrders,'','','','','','','','취소/반품 제외']);
  rows.push(['요약','총매출액',salesAgg.totalGrossSales,'','','','','','','','취소/반품 포함']);
  rows.push(['요약','취소/반품매출',salesAgg.totalCancelSales,'','','','','','','','취소/반품 판정 row']);
  rows.push(['요약','순수매출액',salesAgg.totalNetSales,'','','','','','','','취소/반품 제외']);
  rows.push(['요약','실제정산금액',salesAgg.totalSettlement,'','','','','','','','정산예정금액 / 취소제외']);
  rows.push(['요약','마켓수수료/비용',salesAgg.totalMarketFee,'','','','','','','','순수매출액 - 실제정산금액']);
  rows.push(['요약','마켓수수료율',salesAgg.totalMarketFeeRate,'','','','','','','','마켓수수료 ÷ 순수매출액']);
  rows.push(['요약','매입금액',salesAgg.totalPurchase,'','','','','','','',salesAgg.hasPurchaseColumn ? '취소제외 매입/원가 컬럼 기준' : '매입/원가 컬럼 미확인']);
  rows.push(['요약','순수익',salesAgg.totalNetProfit,'','','','','','','','실제정산금액 - 매입금액']);
  rows.push(['요약','순수익률',salesAgg.totalNetProfitRate,'','','','','','','','순수익 ÷ 순수매출액']);
  var vatBaseTotal = Math.round(salesAgg.totalNetSales / 1.1);
  rows.push(['요약','부가세 신고 공급대가',salesAgg.totalNetSales,'','','','','','','','순수매출액 기준']);
  rows.push(['요약','부가세 신고 공급가액',vatBaseTotal,'','','','','','','','공급대가 ÷ 1.1']);
  rows.push(['요약','부가세',salesAgg.totalNetSales - vatBaseTotal,'','','','','','','','공급대가 - 공급가액']);
  rows.push(['','','','','','','','','','','']);

  rows.push(['구분','브랜드명','대표검색필터명','계정ID','수집수','매출상품','주문','순수매출액','실제정산금액','순수익률','메모']);
  brandRows.slice().sort(function(a, b) { return (b.netSalesAmount || 0) - (a.netSalesAmount || 0); }).slice(0, 30).forEach(function(b) {
    rows.push(['브랜드TOP', b.brand, b.filterName, b.accountId, b.collectCount, b.salesProductCount, b.orderCount, b.netSalesAmount, b.settlementAmount, b.netProfitRate, '총매출 ' + formatWonText_v613_(b.grossSalesAmount) + ' / 취소 ' + formatWonText_v613_(b.cancelSalesAmount) + ' / 순수익 ' + formatWonText_v613_(b.netProfitAmount)]);
  });
  rows.push(['','','','','','','','','','','']);

  rows.push(['구분','브랜드명','대표검색필터명','계정ID','수집수','매출상품','주문','순수매출액','실제정산금액','다음작업','상품갈이메모']);
  brandRows.slice().sort(function(a, b) { return (b.unsoldCount || 0) - (a.unsoldCount || 0); }).slice(0, 40).forEach(function(b) {
    rows.push(['상품갈이', b.brand, b.filterName, b.accountId, b.collectCount, b.salesProductCount, b.orderCount, b.netSalesAmount, b.settlementAmount, decideNextRotation_v611_(b), '최근수집 ' + formatShortDate_(b.latestRecentDate) + ' / 미판매 ' + b.unsoldCount + ' / 순매출 기준']);
  });

  replaceSheetValues_v611_(sheet, CONFIG.HEADERS.DASHBOARD, rows);
  formatDashboardSettlementProfit_v612_(sheet);
}

function readSalesSourceRows_v613_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SALES_CLEAN) || ss.getSheetByName(CONFIG.SHEETS.SALES_IN);
  if (!sheet || sheet.getLastRow() < 2) return { rows: [], sourceSheet: '', hasPurchaseColumn: false };
  var values = sheet.getDataRange().getValues();
  var header = values[0].map(function(h) { return String(h || '').trim(); });
  var rows = [];
  var hasPurchase = header.some(function(h) { return ['매입금액','구매금액','매입가','구매가','원가','상품매입금액','매입금액합계'].indexOf(h) >= 0; });
  for (var r = 1; r < values.length; r++) {
    var obj = {};
    var blank = true;
    for (var c = 0; c < header.length; c++) {
      if (!header[c]) continue;
      obj[header[c]] = values[r][c];
      if (values[r][c] !== '' && values[r][c] != null) blank = false;
    }
    if (!blank) rows.push(obj);
  }
  return { rows: rows, sourceSheet: sheet.getName(), hasPurchaseColumn: hasPurchase };
}

function groupVatDetails_v613_(details, keyFn, initFn) {
  var map = {};
  (details || []).forEach(function(d) {
    var key = keyFn(d);
    if (!map[key]) {
      map[key] = initFn(d);
      map[key].qty = 0;
      map[key].netSales = 0;
      map[key].settlement = 0;
      map[key].marketFee = 0;
      map[key].purchase = 0;
      map[key].netProfit = 0;
      map[key].rows = [];
    }
    var g = map[key];
    g.qty += d.qty || 0;
    g.netSales += d.netSales || 0;
    g.settlement += d.settlement || 0;
    g.marketFee += d.marketFee || 0;
    g.purchase += d.purchase || 0;
    g.netProfit += d.netProfit || 0;
    g.rows.push(d);
  });
  return map;
}

function writeVatSummarySheet_v613_(sheetName, headers, rows, moneyCols, numberCols) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  replaceSheetValues_v611_(sheet, headers, rows);
  formatMoneyColumns_v611_(sheet, moneyCols || []);
  formatNumberColumns_v611_(sheet, numberCols || []);
}

function sortVatRows_v613_(a, b) {
  var ak = String(a[0] || '') + '-' + String(a[1] || '') + '-' + String(a[2] || '') + '-' + String(a[3] || '');
  var bk = String(b[0] || '') + '-' + String(b[1] || '') + '-' + String(b[2] || '') + '-' + String(b[3] || '');
  return ak.localeCompare(bk);
}

function datePartsFromValue_v613_(value, fallbackMonth) {
  var out = { year: '', month: '', day: '', dateText: '' };
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    out.year = Utilities.formatDate(value, CONFIG.TIMEZONE, 'yyyy');
    out.month = Utilities.formatDate(value, CONFIG.TIMEZONE, 'M');
    out.day = Utilities.formatDate(value, CONFIG.TIMEZONE, 'd');
    out.dateText = out.year + '-' + ('0' + out.month).slice(-2) + '-' + ('0' + out.day).slice(-2);
    return out;
  }
  var raw = String(value || '').trim();
  var m = raw.match(/(20\d{2})[-.\/](\d{1,2})[-.\/](\d{1,2})/);
  if (m) {
    out.year = m[1];
    out.month = String(Number(m[2]));
    out.day = String(Number(m[3]));
    out.dateText = out.year + '-' + ('0' + out.month).slice(-2) + '-' + ('0' + out.day).slice(-2);
    return out;
  }
  var fm = String(fallbackMonth || '').trim().match(/(20\d{2})[-.\/](\d{1,2})/);
  if (fm) {
    out.year = fm[1];
    out.month = String(Number(fm[2]));
    out.day = '';
    out.dateText = out.year + '-' + ('0' + out.month).slice(-2);
  }
  return out;
}

function isCancelRow_v613_(obj) {
  var text = [
    getFirstValue_v613_(obj, ['취소여부']),
    getFirstValue_v613_(obj, ['마켓주문상태']),
    getFirstValue_v613_(obj, ['더망고주문상태']),
    getFirstValue_v613_(obj, ['주문상태']),
    getFirstValue_v613_(obj, ['유효여부']),
    getFirstValue_v613_(obj, ['분석대상여부'])
  ].join(' ');
  if (/취소아님|정상|Y\s*$/i.test(text) && !/취소|반품|환불|교환|cancel|return|refund/i.test(text)) return false;
  return /취소|반품|환불|교환|cancel|return|refund|분석제외|N/i.test(text);
}

function getFirstValue_v613_(obj, names) {
  for (var i = 0; i < names.length; i++) {
    if (Object.prototype.hasOwnProperty.call(obj, names[i]) && obj[names[i]] !== '' && obj[names[i]] != null) return obj[names[i]];
  }
  return '';
}

function canonicalBrand_v613_(brandName) {
  return String(brandName || '').trim().replace(/^\d+_?/, '').replace(/_\d+$/, '').trim();
}

function normalizeBrandKey_v613_(brandName) {
  return canonicalBrand_v613_(brandName).toLowerCase().replace(/\s+/g, '').replace(/[\[\]\(\)\{\}\-_]/g, '').trim();
}

function formatWonText_v613_(v) {
  return '₩' + String(Math.round(Number(v || 0))).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}
