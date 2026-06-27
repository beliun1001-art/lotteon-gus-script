/**
 * LOTTEON v6.11 brand-based operation patch
 *
 * 목적:
 * 1) 마켓 전송수/쿠팡전송수 기준 제거
 * 2) 대시보드는 검색필터명이 아니라 브랜드명 기준으로 매출 합산
 * 3) 필터별_상품수의 API_totalCount는 '수집수'로만 사용
 * 4) 쿠팡재전송_로그의 검색필터명/브랜드명/수집수/1단계 추가수집일을 필터별_상품수 기준으로 최신화
 * 5) 매출데이터 기준 브랜드별 마진율, 부가세 신고자료 시트 생성
 *
 * 주의:
 * - 매입금액/원가 컬럼이 매출데이터_정리 또는 매출데이터_붙여넣기에 없으면 마진율은 정산예정금액 기준 참고값으로 표시합니다.
 * - 기존 원본 시트는 삭제하지 않고, 출력 시트와 대시보드 표시를 새 운영 기준으로 재작성합니다.
 */

var LOTTEON_PATCH_V611_BRAND_DASHBOARD_LOG_MARGIN_VAT_LOADED = true;

var __baseBuildDashboard_v611 = typeof buildDashboard_ === 'function' ? buildDashboard_ : null;
var __baseRefreshDashboardFastOnly_v611 = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;
var __baseRunPendingChangesApproval_v611 = typeof runPendingChangesApproval === 'function' ? runPendingChangesApproval : null;
var __baseRefreshCoreSummaryAndDashboard_v611 = typeof refreshCoreSummaryAndDashboardWithRetransmitLogDates === 'function' ? refreshCoreSummaryAndDashboardWithRetransmitLogDates : null;
var __baseCreateRetransmitLogSheet_v611 = typeof createRetransmitLogSheet === 'function' ? createRetransmitLogSheet : null;

if (__baseBuildDashboard_v611) {
  buildDashboard_ = function() {
    var result = __baseBuildDashboard_v611.apply(this, arguments);
    rebuildBrandBasedOperationViews_v611_();
    return result;
  };
}

refreshDashboardFastOnly = function() {
  var result = __baseRefreshDashboardFastOnly_v611 ? __baseRefreshDashboardFastOnly_v611.apply(this, arguments) : null;
  rebuildBrandBasedOperationViews_v611_();
  return result;
};

runPendingChangesApproval = function() {
  var result = __baseRunPendingChangesApproval_v611 ? __baseRunPendingChangesApproval_v611.apply(this, arguments) : null;
  rebuildBrandBasedOperationViews_v611_();
  return result;
};

refreshCoreSummaryAndDashboardWithRetransmitLogDates = function() {
  var result = __baseRefreshCoreSummaryAndDashboard_v611 ? __baseRefreshCoreSummaryAndDashboard_v611.apply(this, arguments) : null;
  rebuildBrandBasedOperationViews_v611_();
  return result;
};

createRetransmitLogSheet = function() {
  var result = __baseCreateRetransmitLogSheet_v611 ? __baseCreateRetransmitLogSheet_v611.apply(this, arguments) : null;
  patchRetransmitLogFromFilterSheet_v611_();
  buildMarginAndVatSheetsFromSales_v611_();
  return result;
};

function rebuildBrandBasedOperationViews_v611_() {
  var salesAgg = aggregateSalesByBrand_v611_();
  var filterAgg = aggregateFiltersByBrand_v611_();
  patchRetransmitLogFromFilterSheet_v611_(filterAgg, salesAgg);
  buildMarginAndVatSheetsFromSales_v611_(salesAgg);
  rebuildDashboardBrandBased_v611_(salesAgg, filterAgg);
  SpreadsheetApp.flush();
}

// -----------------------------------------------------------------------------
// 브랜드 기준 필터 집계
// -----------------------------------------------------------------------------

function aggregateFiltersByBrand_v611_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.FILTERS);
  var result = { byBrand: {}, rows: [], totalCollectCount: 0 };
  if (!sheet || sheet.getLastRow() < 2) return result;

  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var colFilter = findHeaderIndex_v611_(header, ['검색필터명']);
  var colBrand = findHeaderIndex_v611_(header, ['브랜드명']);
  var colAccount = findHeaderIndex_v611_(header, ['쿠팡계정ID']);
  var colTotal = findHeaderIndex_v611_(header, ['API_totalCount', 'APItotalCount']);
  var colRecent = findHeaderIndex_v611_(header, ['API_최근수집일자']);
  var colCreate = findHeaderIndex_v611_(header, ['API_필터생성일']);
  if (colFilter < 0) return result;

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var filterName = String(row[colFilter] || '').trim();
    if (!filterName || filterName.indexOf(CONFIG.FILTER_PREFIX) !== 0) continue;

    var brandRaw = colBrand >= 0 ? String(row[colBrand] || '').trim() : '';
    var brand = canonicalBrand_v611_(brandRaw || brandFromFilterName_v611_(filterName));
    if (!brand) continue;

    var accountId = colAccount >= 0 ? String(row[colAccount] || '').trim() : accountIdFromFilterName_v611_(filterName);
    var total = colTotal >= 0 ? toNumber_(row[colTotal]) : 0;
    var recent = colRecent >= 0 ? normalizeDateText_(row[colRecent]) : '';
    var create = colCreate >= 0 ? normalizeDateText_(row[colCreate]) : '';
    var item = {
      filterName: filterName,
      brand: brand,
      key: normalizeBrandKey_v611_(brand),
      accountId: accountId,
      collectCount: total,
      recentDate: recent,
      createDate: create,
      managedScore: managedFilterScore_v611_(filterName)
    };
    result.rows.push(item);
    result.totalCollectCount += total;

    var key = item.key;
    if (!result.byBrand[key]) {
      result.byBrand[key] = {
        brand: brand,
        key: key,
        collectCount: 0,
        filters: [],
        accountIds: {},
        latestRecentDate: '',
        earliestCreateDate: '',
        representativeFilterName: '',
        representativeAccountId: ''
      };
    }
    var b = result.byBrand[key];
    b.collectCount += total;
    b.filters.push(item);
    if (accountId) b.accountIds[accountId] = true;
    if (recent && (!b.latestRecentDate || recent > b.latestRecentDate)) b.latestRecentDate = recent;
    if (create && (!b.earliestCreateDate || create < b.earliestCreateDate)) b.earliestCreateDate = create;
  }

  Object.keys(result.byBrand).forEach(function(key) {
    var b = result.byBrand[key];
    var rep = chooseRepresentativeFilter_v611_(b.filters);
    b.representativeFilterName = rep ? rep.filterName : '';
    b.representativeAccountId = rep ? rep.accountId : Object.keys(b.accountIds)[0] || '';
  });

  return result;
}

function chooseRepresentativeFilter_v611_(list) {
  return (list || []).slice().sort(function(a, b) {
    if ((b.collectCount || 0) !== (a.collectCount || 0)) return (b.collectCount || 0) - (a.collectCount || 0);
    if ((b.managedScore || 0) !== (a.managedScore || 0)) return (b.managedScore || 0) - (a.managedScore || 0);
    return String(a.filterName || '').localeCompare(String(b.filterName || ''));
  })[0] || null;
}

// -----------------------------------------------------------------------------
// 매출 브랜드 집계
// -----------------------------------------------------------------------------

function aggregateSalesByBrand_v611_() {
  var table = readSalesSourceRows_v611_();
  var result = {
    byBrand: {},
    rows: table.rows,
    totalSales: 0,
    totalSettlement: 0,
    totalPurchase: 0,
    totalQty: 0,
    totalOrders: 0,
    totalProducts: 0,
    cancelQty: 0,
    cancelSales: 0,
    hasPurchaseColumn: table.hasPurchaseColumn,
    sourceSheet: table.sourceSheet
  };

  var orderSet = {};
  var productSet = {};

  (table.rows || []).forEach(function(rowObj) {
    var brand = canonicalBrand_v611_(getFirstValue_v611_(rowObj, ['브랜드명_매칭','브랜드명','브랜드명_원본','브랜드','매출브랜드명']));
    if (!brand) brand = '브랜드미확인';
    var key = normalizeBrandKey_v611_(brand);
    var salesAmount = toNumber_(getFirstValue_v611_(rowObj, ['결제금액합계','결제금액','매출액','판매금액','총결제금액','상품금액']));
    var settlement = toNumber_(getFirstValue_v611_(rowObj, ['정산예정금액','정산금액','정산액']));
    var purchase = toNumber_(getFirstValue_v611_(rowObj, ['매입금액','구매금액','매입가','구매가','원가','상품매입금액','매입금액합계']));
    var qty = toNumber_(getFirstValue_v611_(rowObj, ['결제수량','수량','판매수량']));
    if (!qty && salesAmount) qty = 1;
    var orderNo = String(getFirstValue_v611_(rowObj, ['마켓주문번호','주문번호','주문ID']) || '').trim();
    var productNo = String(getFirstValue_v611_(rowObj, ['사이트상품번호','마켓상품번호','상품번호','상품ID']) || '').trim();
    var month = String(getFirstValue_v611_(rowObj, ['주문월','월']) || '').trim() || monthFromDateText_v611_(getFirstValue_v611_(rowObj, ['주문일시','주문일자','결제일시','결제일']));
    var cancel = isCancelRow_v611_(rowObj);

    if (!result.byBrand[key]) {
      result.byBrand[key] = {
        brand: brand,
        key: key,
        salesAmount: 0,
        settlementAmount: 0,
        purchaseAmount: 0,
        quantity: 0,
        orderSet: {},
        productSet: {},
        orderCount: 0,
        salesProductCount: 0,
        cancelQty: 0,
        cancelSales: 0,
        netSalesForVat: 0,
        netVatBase: 0,
        netVat: 0,
        months: {}
      };
    }
    var b = result.byBrand[key];
    b.salesAmount += salesAmount;
    b.settlementAmount += settlement;
    b.purchaseAmount += purchase;
    b.quantity += qty;
    if (orderNo) { b.orderSet[orderNo] = true; orderSet[orderNo] = true; }
    if (productNo) { b.productSet[productNo] = true; productSet[productNo] = true; }
    if (cancel) {
      b.cancelQty += qty;
      b.cancelSales += salesAmount;
      result.cancelQty += qty;
      result.cancelSales += salesAmount;
    } else {
      b.netSalesForVat += salesAmount;
    }

    if (!b.months[month]) b.months[month] = { month: month, sales: 0, settlement: 0, purchase: 0, qty: 0, orderSet: {}, netSalesForVat: 0 };
    var m = b.months[month];
    m.sales += salesAmount;
    m.settlement += settlement;
    m.purchase += purchase;
    m.qty += qty;
    if (orderNo) m.orderSet[orderNo] = true;
    if (!cancel) m.netSalesForVat += salesAmount;

    result.totalSales += salesAmount;
    result.totalSettlement += settlement;
    result.totalPurchase += purchase;
    result.totalQty += qty;
  });

  Object.keys(result.byBrand).forEach(function(key) {
    var b = result.byBrand[key];
    b.orderCount = Object.keys(b.orderSet).length || (b.salesAmount ? 1 : 0);
    b.salesProductCount = Object.keys(b.productSet).length;
    b.marginAmount = b.settlementAmount - b.purchaseAmount;
    b.marginRate = b.salesAmount ? b.marginAmount / b.salesAmount : 0;
    b.netVatBase = Math.round(b.netSalesForVat / 1.1);
    b.netVat = b.netSalesForVat - b.netVatBase;
  });

  result.totalOrders = Object.keys(orderSet).length;
  result.totalProducts = Object.keys(productSet).length;
  return result;
}

function readSalesSourceRows_v611_() {
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

// -----------------------------------------------------------------------------
// 쿠팡재전송_로그 최신화: 필터명/수집수/1단계 추가수집일
// -----------------------------------------------------------------------------

function patchRetransmitLogFromFilterSheet_v611_(filterAgg, salesAgg) {
  filterAgg = filterAgg || aggregateFiltersByBrand_v611_();
  salesAgg = salesAgg || aggregateSalesByBrand_v611_();

  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.RETRANSMIT_LOG);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.RETRANSMIT_LOG);

  var headers = ['검색필터명','브랜드명','쿠팡\n계정ID','최초\n전송일','누적\n매출액','월평균\n매출액','작업\n유형','1단계_\n추가수집일','2단계_\n재전송일','3단계_\n정리재수집일','최종_\n더망고수집수','추가수집\n상품수','작업\n사유','비고'];
  var rows = [];

  Object.keys(filterAgg.byBrand).sort(function(a, b) {
    return String(filterAgg.byBrand[a].brand).localeCompare(String(filterAgg.byBrand[b].brand));
  }).forEach(function(key) {
    var f = filterAgg.byBrand[key];
    var s = salesAgg.byBrand[key] || {};
    var sales = s.salesAmount || 0;
    var avg = sales ? Math.round(sales / Math.max(1, countDistinctMonths_v611_(s.months))) : 0;
    var recent = f.latestRecentDate || '';
    var create = f.earliestCreateDate || '';
    var collect = f.collectCount || 0;
    var soldProducts = s.salesProductCount || 0;
    var unsold = Math.max(0, collect - soldProducts);
    var actionType = decideRetransmitActionType_v611_(collect, soldProducts, sales, recent, create);
    rows.push([
      f.representativeFilterName || '',
      f.brand,
      f.representativeAccountId || '',
      formatShortDate_(create),
      sales,
      avg,
      actionType,
      formatShortDate_(recent),
      '',
      '',
      collect,
      unsold,
      actionType,
      'v6.11: 필터별_상품수 API_최근수집일자 기준 / 마켓전송수 미사용'
    ]);
  });

  replaceSheetValues_v611_(sheet, headers, rows);
  formatRetransmitLog_v611_(sheet);
  return { rows: rows.length };
}

function decideRetransmitActionType_v611_(collect, soldProducts, sales, recent, create) {
  if (!collect) return '수집확인';
  if (!sales && collect >= 300) return '압축검토';
  if (recent) return '추가수집관리';
  return '관찰';
}

// -----------------------------------------------------------------------------
// 대시보드: 브랜드 기준 재작성, 마켓 전송수 정보 제거
// -----------------------------------------------------------------------------

function rebuildDashboardBrandBased_v611_(salesAgg, filterAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD) || ss.insertSheet(CONFIG.SHEETS.DASHBOARD);
  salesAgg = salesAgg || aggregateSalesByBrand_v611_();
  filterAgg = filterAgg || aggregateFiltersByBrand_v611_();

  var brandRows = buildBrandRowsForDashboard_v611_(salesAgg, filterAgg);
  var nowText = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  var rows = [];

  rows.push(['요약','갱신기준','매출데이터 브랜드 기준 + 필터별_상품수 수집일 기준','','','','','','','','v6.11: 검색필터별 분리 제거']);
  rows.push(['요약','갱신시각',nowText,'','','','','','','','Asia/Seoul']);
  rows.push(['요약','분석브랜드수',brandRows.length,'','','','','','','','브랜드명 기준']);
  rows.push(['요약','수집상품수',filterAgg.totalCollectCount,'','','','','','','','필터별_상품수 API_totalCount 총합']);
  rows.push(['요약','매출상품수',salesAgg.totalProducts,'','','','','','','','사이트/마켓 상품번호 기준']);
  rows.push(['요약','주문건수',salesAgg.totalOrders,'','','','','','','','브랜드 합산']);
  rows.push(['요약','매출액',salesAgg.totalSales,'','','','','','','','매출데이터 기준']);
  rows.push(['요약','정산예정금액',salesAgg.totalSettlement,'','','','','','','','매출데이터 기준']);
  rows.push(['요약','매입금액',salesAgg.totalPurchase,'','','','','','','',salesAgg.hasPurchaseColumn ? '매출데이터 매입/원가 컬럼 기준' : '매입/원가 컬럼 미확인']);
  rows.push(['요약','추정마진액',salesAgg.totalSettlement - salesAgg.totalPurchase,'','','','','','','','정산예정금액 - 매입금액']);
  rows.push(['요약','추정마진율',salesAgg.totalSales ? (salesAgg.totalSettlement - salesAgg.totalPurchase) / salesAgg.totalSales : 0,'','','','','','','','매출액 대비']);
  rows.push(['요약','부가세 신고 공급대가',sumNetVatSales_v611_(salesAgg),'','','','','','','','취소/반품 제외 매출 기준']);
  var vatBase = Math.round(sumNetVatSales_v611_(salesAgg) / 1.1);
  rows.push(['요약','부가세 신고 공급가액',vatBase,'','','','','','','','공급대가 ÷ 1.1']);
  rows.push(['요약','부가세',sumNetVatSales_v611_(salesAgg) - vatBase,'','','','','','','','공급대가 - 공급가액']);
  rows.push(['요약','취소/반품 수량',salesAgg.cancelQty,'','','','','','','','정보성']);
  rows.push(['요약','취소/반품 매출',salesAgg.cancelSales,'','','','','','','','정보성']);
  rows.push(['','','','','','','','','','','']);

  rows.push(['구분','브랜드명','대표검색필터명','계정ID','수집수','매출상품','주문','매출','정산예정','마진율','메모']);
  brandRows.slice().sort(function(a, b) { return (b.salesAmount || 0) - (a.salesAmount || 0); }).slice(0, 30).forEach(function(b) {
    rows.push(['브랜드TOP', b.brand, b.filterName, b.accountId, b.collectCount, b.salesProductCount, b.orderCount, b.salesAmount, b.settlementAmount, b.marginRate, '브랜드 기준 합산 / 마켓전송수 미사용']);
  });
  rows.push(['','','','','','','','','','','']);

  rows.push(['구분','브랜드명','대표검색필터명','계정ID','수집수','매출상품','주문','매출','정산예정','다음작업','상품갈이메모']);
  brandRows.slice().sort(function(a, b) { return (b.unsoldCount || 0) - (a.unsoldCount || 0); }).slice(0, 40).forEach(function(b) {
    var next = decideNextRotation_v611_(b);
    rows.push(['상품갈이', b.brand, b.filterName, b.accountId, b.collectCount, b.salesProductCount, b.orderCount, b.salesAmount, b.settlementAmount, next, '최근수집 ' + formatShortDate_(b.latestRecentDate) + ' / 미판매 ' + b.unsoldCount + ' / 브랜드합산']);
  });
  rows.push(['','','','','','','','','','','']);

  rows.push(['구분','브랜드명','대표검색필터명','계정ID','수집수','매출상품','주문','매출','정산예정','액션','메모']);
  brandRows.filter(function(b) { return b.orderCount === 0 || b.salesAmount === 0 || !b.filterName; }).slice(0, 40).forEach(function(b) {
    rows.push(['조치필요', b.brand, b.filterName, b.accountId, b.collectCount, b.salesProductCount, b.orderCount, b.salesAmount, b.settlementAmount, b.filterName ? '관찰/정리' : '필터확인', '브랜드 기준 / 마켓전송수 미사용']);
  });

  replaceSheetValues_v611_(sheet, CONFIG.HEADERS.DASHBOARD, rows);
  formatDashboard_v611_(sheet);
  return { rows: rows.length };
}

function buildBrandRowsForDashboard_v611_(salesAgg, filterAgg) {
  var keys = {};
  Object.keys(salesAgg.byBrand || {}).forEach(function(k) { keys[k] = true; });
  Object.keys(filterAgg.byBrand || {}).forEach(function(k) { keys[k] = true; });
  return Object.keys(keys).map(function(key) {
    var s = salesAgg.byBrand[key] || {};
    var f = filterAgg.byBrand[key] || {};
    var brand = s.brand || f.brand || '브랜드미확인';
    var collect = f.collectCount || 0;
    var sold = s.salesProductCount || 0;
    return {
      brand: brand,
      key: key,
      filterName: f.representativeFilterName || '',
      accountId: f.representativeAccountId || '',
      collectCount: collect,
      salesProductCount: sold,
      orderCount: s.orderCount || 0,
      salesAmount: s.salesAmount || 0,
      settlementAmount: s.settlementAmount || 0,
      purchaseAmount: s.purchaseAmount || 0,
      marginAmount: s.marginAmount || 0,
      marginRate: s.marginRate || 0,
      latestRecentDate: f.latestRecentDate || '',
      earliestCreateDate: f.earliestCreateDate || '',
      unsoldCount: Math.max(0, collect - sold)
    };
  });
}

function decideNextRotation_v611_(b) {
  if (!b.collectCount) return '수집확인';
  if (!b.salesAmount && b.collectCount >= 300) return '압축검토';
  if (b.unsoldCount >= 300) return '추가수집/정리검토';
  return '관찰';
}

// -----------------------------------------------------------------------------
// 브랜드별 마진율 / 부가세 신고자료
// -----------------------------------------------------------------------------

function buildMarginAndVatSheetsFromSales_v611_(salesAgg) {
  salesAgg = salesAgg || aggregateSalesByBrand_v611_();
  buildBrandMarginSheet_v611_(salesAgg);
  buildVatReportSheet_v611_(salesAgg);
}

function buildBrandMarginSheet_v611_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('브랜드별_마진율') || ss.insertSheet('브랜드별_마진율');
  var headers = ['브랜드명','주문건수','판매수량','매출상품수','매출액','정산예정금액','매입금액','추정마진액','마진율','비고'];
  var rows = Object.keys(salesAgg.byBrand || {}).map(function(key) {
    var b = salesAgg.byBrand[key];
    return [
      b.brand,
      b.orderCount,
      b.quantity,
      b.salesProductCount,
      b.salesAmount,
      b.settlementAmount,
      b.purchaseAmount,
      b.marginAmount,
      b.marginRate,
      salesAgg.hasPurchaseColumn ? '매입/원가 컬럼 기준' : '매입/원가 컬럼 미확인: 정산예정금액 기준 참고'
    ];
  }).sort(function(a, b) { return b[4] - a[4]; });
  replaceSheetValues_v611_(sheet, headers, rows);
  formatMoneyColumns_v611_(sheet, [5,6,7,8]);
  formatNumberColumns_v611_(sheet, [2,3,4]);
  if (rows.length) sheet.getRange(2, 9, rows.length, 1).setNumberFormat('0.00%');
}

function buildVatReportSheet_v611_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('부가세_신고자료') || ss.insertSheet('부가세_신고자료');
  var headers = ['주문월','브랜드명','공급대가_신고대상','공급가액','부가세','총매출_취소포함','정산예정금액','주문건수','판매수량','비고'];
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
        Object.keys(m.orderSet || {}).length,
        m.qty,
        '취소/반품 제외 매출 기준. 최종 신고 전 원장 대조 필요'
      ]);
    });
  });
  rows.sort(function(a, b) { return String(a[0]).localeCompare(String(b[0])) || String(a[1]).localeCompare(String(b[1])); });
  replaceSheetValues_v611_(sheet, headers, rows);
  formatMoneyColumns_v611_(sheet, [3,4,5,6,7]);
  formatNumberColumns_v611_(sheet, [8,9]);
}

// -----------------------------------------------------------------------------
// 공통 helper
// -----------------------------------------------------------------------------

function replaceSheetValues_v611_(sheet, headers, rows) {
  sheet.clearContents();
  var all = [headers].concat(rows || []);
  if (all.length && all[0].length) sheet.getRange(1, 1, all.length, all[0].length).setValues(all);
  sheet.setFrozenRows(1);
  try { sheet.autoResizeColumns(1, all[0].length); } catch (e) {}
}

function formatDashboard_v611_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  try {
    sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).setFontFamily('Arial').setFontSize(10);
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
    sheet.getRange(1, 3, lastRow, 7).setNumberFormat('#,##0');
    sheet.getRange(1, 10, lastRow, 1).setNumberFormat('0.00%');
    sheet.getRange(1, 7, lastRow, 1).setNumberFormat('#,##0');
    sheet.getRange(1, 8, lastRow, 2).setNumberFormat('₩#,##0');
    sheet.getRange(1, 3, 20, 1).setNumberFormat('#,##0');
  } catch (e) {}
}

function formatRetransmitLog_v611_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  try {
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
    sheet.getRange(2, 5, lastRow - 1, 2).setNumberFormat('₩#,##0');
    sheet.getRange(2, 11, lastRow - 1, 2).setNumberFormat('#,##0');
  } catch (e) {}
}

function formatMoneyColumns_v611_(sheet, cols) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  (cols || []).forEach(function(col) { try { sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('₩#,##0'); } catch (e) {} });
}

function formatNumberColumns_v611_(sheet, cols) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  (cols || []).forEach(function(col) { try { sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('#,##0'); } catch (e) {} });
}

function findHeaderIndex_v611_(headerRow, candidates) {
  var normalized = (headerRow || []).map(function(h) { return normalizeHeaderKey_v611_(h); });
  for (var i = 0; i < candidates.length; i++) {
    var key = normalizeHeaderKey_v611_(candidates[i]);
    var idx = normalized.indexOf(key);
    if (idx >= 0) return idx;
  }
  return -1;
}

function normalizeHeaderKey_v611_(v) {
  return String(v || '').replace(/\s+/g, '').replace(/\n/g, '').replace(/_/g, '').trim();
}

function getFirstValue_v611_(obj, names) {
  for (var i = 0; i < names.length; i++) {
    if (Object.prototype.hasOwnProperty.call(obj, names[i]) && obj[names[i]] !== '' && obj[names[i]] != null) return obj[names[i]];
  }
  return '';
}

function canonicalBrand_v611_(brandName) {
  return String(brandName || '')
    .trim()
    .replace(/^\d+_?/, '')
    .replace(/_\d+$/, '')
    .trim();
}

function normalizeBrandKey_v611_(brandName) {
  return canonicalBrand_v611_(brandName).toLowerCase().replace(/\s+/g, '').replace(/[\[\]\(\)\{\}\-_]/g, '').trim();
}

function brandFromFilterName_v611_(filterName) {
  var text = String(filterName || '').trim();
  var first = text.indexOf('_');
  if (first < 0) return canonicalBrand_v611_(text);
  var second = text.indexOf('_', first + 1);
  if (second < 0) return canonicalBrand_v611_(text);
  return canonicalBrand_v611_(text.slice(second + 1));
}

function accountIdFromFilterName_v611_(filterName) {
  var text = String(filterName || '').trim();
  if (text.indexOf('롯백_01_') === 0) return 'beliun1021';
  if (text.indexOf('롯백_02_') === 0) return 'beliun1024';
  if (text.indexOf('롯백_03_') === 0) return 'beliun1023';
  if (text.indexOf('롯백_04_') === 0) return 'beliun1024';
  return '';
}

function managedFilterScore_v611_(filterName) {
  var tail = String(filterName || '').split('_').slice(2).join('_');
  var score = 0;
  if (/^\d+_?/.test(tail)) score += 100;
  if (/_\d+$/.test(tail)) score += 50;
  return score;
}

function isCancelRow_v611_(obj) {
  var text = [
    getFirstValue_v611_(obj, ['취소여부']),
    getFirstValue_v611_(obj, ['마켓주문상태']),
    getFirstValue_v611_(obj, ['더망고주문상태']),
    getFirstValue_v611_(obj, ['주문상태'])
  ].join(' ');
  return /Y|취소|반품|환불|교환|cancel|return|refund/i.test(text);
}

function monthFromDateText_v611_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM');
  var m = String(v).match(/(20\d{2})[-.\/](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);
  return '';
}

function countDistinctMonths_v611_(months) {
  var n = Object.keys(months || {}).filter(function(k) { return !!k; }).length;
  return n || 1;
}

function sumNetVatSales_v611_(salesAgg) {
  var total = 0;
  Object.keys(salesAgg.byBrand || {}).forEach(function(key) { total += salesAgg.byBrand[key].netSalesForVat || 0; });
  return total;
}
