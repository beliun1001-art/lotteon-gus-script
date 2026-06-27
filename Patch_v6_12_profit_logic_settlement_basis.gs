/**
 * LOTTEON v6.12 settlement-based profit logic patch
 *
 * 확정 기준:
 * - 정산예정금액 = 실제 정산금액
 * - 마켓수수료/비용 = 매출금액 - 정산예정금액
 * - 순수익 = 정산예정금액 - 매입금액
 * - 순수익률 = 순수익 / 매출금액
 *
 * 목적:
 * - v6.11의 브랜드별 마진율/대시보드 표현을 위 기준으로 명확히 정리합니다.
 * - 부가세 신고자료는 매출액 기준 공급대가/공급가액/부가세 산출을 유지합니다.
 */

var LOTTEON_PATCH_V612_PROFIT_LOGIC_SETTLEMENT_BASIS_LOADED = true;

var __baseAggregateSalesByBrand_v612 = typeof aggregateSalesByBrand_v611_ === 'function' ? aggregateSalesByBrand_v611_ : null;
var __baseRebuildDashboardBrandBased_v612 = typeof rebuildDashboardBrandBased_v611_ === 'function' ? rebuildDashboardBrandBased_v611_ : null;
var __baseBuildBrandMarginSheet_v612 = typeof buildBrandMarginSheet_v611_ === 'function' ? buildBrandMarginSheet_v611_ : null;
var __baseBuildVatReportSheet_v612 = typeof buildVatReportSheet_v611_ === 'function' ? buildVatReportSheet_v611_ : null;

if (__baseAggregateSalesByBrand_v612) {
  aggregateSalesByBrand_v611_ = function() {
    var result = __baseAggregateSalesByBrand_v612.apply(this, arguments);
    return applySettlementProfitMetrics_v612_(result);
  };
}

if (__baseRebuildDashboardBrandBased_v612) {
  rebuildDashboardBrandBased_v611_ = function(salesAgg, filterAgg) {
    salesAgg = applySettlementProfitMetrics_v612_(salesAgg || aggregateSalesByBrand_v611_());
    return rebuildDashboardBrandBasedSettlementProfit_v612_(salesAgg, filterAgg || aggregateFiltersByBrand_v611_());
  };
}

if (__baseBuildBrandMarginSheet_v612) {
  buildBrandMarginSheet_v611_ = function(salesAgg) {
    return buildBrandProfitSheetSettlementBasis_v612_(applySettlementProfitMetrics_v612_(salesAgg || aggregateSalesByBrand_v611_()));
  };
}

if (__baseBuildVatReportSheet_v612) {
  buildVatReportSheet_v611_ = function(salesAgg) {
    salesAgg = applySettlementProfitMetrics_v612_(salesAgg || aggregateSalesByBrand_v611_());
    return buildVatReportSheetSettlementBasis_v612_(salesAgg);
  };
}

function applySettlementProfitMetrics_v612_(salesAgg) {
  salesAgg = salesAgg || {};
  salesAgg.totalSales = salesAgg.totalSales || 0;
  salesAgg.totalSettlement = salesAgg.totalSettlement || 0;
  salesAgg.totalPurchase = salesAgg.totalPurchase || 0;
  salesAgg.totalMarketFee = salesAgg.totalSales - salesAgg.totalSettlement;
  salesAgg.totalNetProfit = salesAgg.totalSettlement - salesAgg.totalPurchase;
  salesAgg.totalMarketFeeRate = salesAgg.totalSales ? salesAgg.totalMarketFee / salesAgg.totalSales : 0;
  salesAgg.totalNetProfitRate = salesAgg.totalSales ? salesAgg.totalNetProfit / salesAgg.totalSales : 0;

  Object.keys(salesAgg.byBrand || {}).forEach(function(key) {
    var b = salesAgg.byBrand[key] || {};
    b.marketFeeAmount = (b.salesAmount || 0) - (b.settlementAmount || 0);
    b.marketFeeRate = b.salesAmount ? b.marketFeeAmount / b.salesAmount : 0;
    b.netProfitAmount = (b.settlementAmount || 0) - (b.purchaseAmount || 0);
    b.netProfitRate = b.salesAmount ? b.netProfitAmount / b.salesAmount : 0;

    // v6.11 호환 컬럼명 유지: 기존 marginAmount/marginRate는 순수익/순수익률로 재정의
    b.marginAmount = b.netProfitAmount;
    b.marginRate = b.netProfitRate;

    Object.keys(b.months || {}).forEach(function(month) {
      var m = b.months[month] || {};
      m.marketFeeAmount = (m.sales || 0) - (m.settlement || 0);
      m.netProfitAmount = (m.settlement || 0) - (m.purchase || 0);
      m.marketFeeRate = m.sales ? m.marketFeeAmount / m.sales : 0;
      m.netProfitRate = m.sales ? m.netProfitAmount / m.sales : 0;
    });
  });

  return salesAgg;
}

function rebuildDashboardBrandBasedSettlementProfit_v612_(salesAgg, filterAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD) || ss.insertSheet(CONFIG.SHEETS.DASHBOARD);
  salesAgg = applySettlementProfitMetrics_v612_(salesAgg || aggregateSalesByBrand_v611_());
  filterAgg = filterAgg || aggregateFiltersByBrand_v611_();

  var brandRows = buildBrandRowsForDashboard_v611_(salesAgg, filterAgg);
  brandRows.forEach(function(b) {
    b.marketFeeAmount = (b.salesAmount || 0) - (b.settlementAmount || 0);
    b.marketFeeRate = b.salesAmount ? b.marketFeeAmount / b.salesAmount : 0;
    b.netProfitAmount = (b.settlementAmount || 0) - (b.purchaseAmount || 0);
    b.netProfitRate = b.salesAmount ? b.netProfitAmount / b.salesAmount : 0;
    b.marginAmount = b.netProfitAmount;
    b.marginRate = b.netProfitRate;
  });

  var nowText = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  var rows = [];

  rows.push(['요약','갱신기준','브랜드 기준 + 정산예정금액 실제정산 기준','','','','','','','','v6.12: 수수료/순수익 계산식 반영']);
  rows.push(['요약','갱신시각',nowText,'','','','','','','','Asia/Seoul']);
  rows.push(['요약','분석브랜드수',brandRows.length,'','','','','','','','브랜드명 기준']);
  rows.push(['요약','수집상품수',filterAgg.totalCollectCount,'','','','','','','','필터별_상품수 API_totalCount 총합']);
  rows.push(['요약','매출상품수',salesAgg.totalProducts,'','','','','','','','사이트/마켓 상품번호 기준']);
  rows.push(['요약','주문건수',salesAgg.totalOrders,'','','','','','','','브랜드 합산']);
  rows.push(['요약','매출금액',salesAgg.totalSales,'','','','','','','','매출데이터 결제금액합계']);
  rows.push(['요약','실제정산금액',salesAgg.totalSettlement,'','','','','','','','정산예정금액']);
  rows.push(['요약','마켓수수료/비용',salesAgg.totalMarketFee,'','','','','','','','매출금액 - 실제정산금액']);
  rows.push(['요약','마켓수수료율',salesAgg.totalMarketFeeRate,'','','','','','','','마켓수수료 ÷ 매출금액']);
  rows.push(['요약','매입금액',salesAgg.totalPurchase,'','','','','','','',salesAgg.hasPurchaseColumn ? '매출데이터 매입/원가 컬럼 기준' : '매입/원가 컬럼 미확인']);
  rows.push(['요약','순수익',salesAgg.totalNetProfit,'','','','','','','','실제정산금액 - 매입금액']);
  rows.push(['요약','순수익률',salesAgg.totalNetProfitRate,'','','','','','','','순수익 ÷ 매출금액']);
  rows.push(['요약','부가세 신고 공급대가',sumNetVatSales_v611_(salesAgg),'','','','','','','','취소/반품 제외 매출 기준']);
  var vatBase = Math.round(sumNetVatSales_v611_(salesAgg) / 1.1);
  rows.push(['요약','부가세 신고 공급가액',vatBase,'','','','','','','','공급대가 ÷ 1.1']);
  rows.push(['요약','부가세',sumNetVatSales_v611_(salesAgg) - vatBase,'','','','','','','','공급대가 - 공급가액']);
  rows.push(['요약','취소/반품 수량',salesAgg.cancelQty,'','','','','','','','정보성']);
  rows.push(['요약','취소/반품 매출',salesAgg.cancelSales,'','','','','','','','정보성']);
  rows.push(['','','','','','','','','','','']);

  rows.push(['구분','브랜드명','대표검색필터명','계정ID','수집수','매출상품','주문','매출금액','실제정산금액','순수익률','메모']);
  brandRows.slice().sort(function(a, b) { return (b.salesAmount || 0) - (a.salesAmount || 0); }).slice(0, 30).forEach(function(b) {
    rows.push(['브랜드TOP', b.brand, b.filterName, b.accountId, b.collectCount, b.salesProductCount, b.orderCount, b.salesAmount, b.settlementAmount, b.netProfitRate, '마켓수수료 ' + formatWonText_v612_(b.marketFeeAmount) + ' / 순수익 ' + formatWonText_v612_(b.netProfitAmount)]);
  });
  rows.push(['','','','','','','','','','','']);

  rows.push(['구분','브랜드명','대표검색필터명','계정ID','수집수','매출상품','주문','매출금액','실제정산금액','다음작업','상품갈이메모']);
  brandRows.slice().sort(function(a, b) { return (b.unsoldCount || 0) - (a.unsoldCount || 0); }).slice(0, 40).forEach(function(b) {
    var next = decideNextRotation_v611_(b);
    rows.push(['상품갈이', b.brand, b.filterName, b.accountId, b.collectCount, b.salesProductCount, b.orderCount, b.salesAmount, b.settlementAmount, next, '최근수집 ' + formatShortDate_(b.latestRecentDate) + ' / 미판매 ' + b.unsoldCount + ' / 브랜드합산']);
  });
  rows.push(['','','','','','','','','','','']);

  rows.push(['구분','브랜드명','대표검색필터명','계정ID','수집수','매출상품','주문','매출금액','실제정산금액','액션','메모']);
  brandRows.filter(function(b) { return b.orderCount === 0 || b.salesAmount === 0 || !b.filterName; }).slice(0, 40).forEach(function(b) {
    rows.push(['조치필요', b.brand, b.filterName, b.accountId, b.collectCount, b.salesProductCount, b.orderCount, b.salesAmount, b.settlementAmount, b.filterName ? '관찰/정리' : '필터확인', '브랜드 기준 / 마켓전송수 미사용']);
  });

  replaceSheetValues_v611_(sheet, CONFIG.HEADERS.DASHBOARD, rows);
  formatDashboardSettlementProfit_v612_(sheet);
  return { rows: rows.length };
}

function buildBrandProfitSheetSettlementBasis_v612_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('브랜드별_마진율') || ss.insertSheet('브랜드별_마진율');
  var headers = ['브랜드명','주문건수','판매수량','매출상품수','매출금액','실제정산금액','마켓수수료/비용','마켓수수료율','매입금액','순수익','순수익률','비고'];
  var rows = Object.keys(salesAgg.byBrand || {}).map(function(key) {
    var b = salesAgg.byBrand[key];
    return [
      b.brand,
      b.orderCount,
      b.quantity,
      b.salesProductCount,
      b.salesAmount,
      b.settlementAmount,
      b.marketFeeAmount,
      b.marketFeeRate,
      b.purchaseAmount,
      b.netProfitAmount,
      b.netProfitRate,
      salesAgg.hasPurchaseColumn ? '정산예정금액=실제정산 / 순수익=정산예정금액-매입금액' : '매입/원가 컬럼 미확인: 순수익 검증 필요'
    ];
  }).sort(function(a, b) { return b[4] - a[4]; });
  replaceSheetValues_v611_(sheet, headers, rows);
  formatMoneyColumns_v611_(sheet, [5,6,7,9,10]);
  formatNumberColumns_v611_(sheet, [2,3,4]);
  if (rows.length) {
    sheet.getRange(2, 8, rows.length, 1).setNumberFormat('0.00%');
    sheet.getRange(2, 11, rows.length, 1).setNumberFormat('0.00%');
  }
}

function buildVatReportSheetSettlementBasis_v612_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('부가세_신고자료') || ss.insertSheet('부가세_신고자료');
  var headers = ['주문월','브랜드명','공급대가_신고대상','공급가액','부가세','총매출_취소포함','실제정산금액','마켓수수료/비용','매입금액','순수익','주문건수','판매수량','비고'];
  var rows = [];
  Object.keys(salesAgg.byBrand || {}).forEach(function(key) {
    var b = salesAgg.byBrand[key];
    Object.keys(b.months || {}).forEach(function(month) {
      var m = b.months[month];
      var supplyTotal = m.netSalesForVat || 0;
      var supplyBase = Math.round(supplyTotal / 1.1);
      var vat = supplyTotal - supplyBase;
      rows.push([
        month,
        b.brand,
        supplyTotal,
        supplyBase,
        vat,
        m.sales,
        m.settlement,
        m.marketFeeAmount,
        m.purchase,
        m.netProfitAmount,
        Object.keys(m.orderSet || {}).length,
        m.qty,
        '부가세는 매출 공급대가 기준 / 정산예정금액은 실제정산 참고'
      ]);
    });
  });
  rows.sort(function(a, b) { return String(a[0]).localeCompare(String(b[0])) || String(a[1]).localeCompare(String(b[1])); });
  replaceSheetValues_v611_(sheet, headers, rows);
  formatMoneyColumns_v611_(sheet, [3,4,5,6,7,8,9,10]);
  formatNumberColumns_v611_(sheet, [11,12]);
}

function formatDashboardSettlementProfit_v612_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  try {
    sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).setFontFamily('Arial').setFontSize(10);
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
    sheet.getRange(1, 3, lastRow, 8).setNumberFormat('#,##0');
    sheet.getRange(1, 8, lastRow, 2).setNumberFormat('₩#,##0');
    sheet.getRange(1, 10, lastRow, 1).setNumberFormat('0.00%');
  } catch (e) {}
}

function formatWonText_v612_(v) {
  return '₩' + String(Math.round(Number(v || 0))).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}
