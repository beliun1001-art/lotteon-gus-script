/**
 * LOTTEON v6.14 purchase amount and year format fix patch
 *
 * 수정 내용:
 * 1) 부가세 관련 시트의 년/월/일 컬럼 서식을 #,##0이 아니라 0으로 고정합니다.
 *    - 2,026 -> 2026
 * 2) 매입금액 컬럼 탐지 범위를 넓힙니다.
 *    - 매출데이터_정리에 매입금액이 없고 매출데이터_붙여넣기에만 있을 때도 주문/상품 기준으로 보강합니다.
 * 3) 매입금액이 잡히면 브랜드별_마진율/부가세 상세/상품별/고객별/주문번호별 순수익을 다시 계산합니다.
 */

var LOTTEON_PATCH_V614_PURCHASE_AND_YEAR_FORMAT_FIX_LOADED = true;

var __baseAggregateSalesByBrand_v614 = typeof aggregateSalesByBrand_v611_ === 'function' ? aggregateSalesByBrand_v611_ : null;
var __baseRebuildDashboardBrandBased_v614 = typeof rebuildDashboardBrandBased_v611_ === 'function' ? rebuildDashboardBrandBased_v611_ : null;
var __baseBuildBrandMarginSheet_v614 = typeof buildBrandMarginSheet_v611_ === 'function' ? buildBrandMarginSheet_v611_ : null;
var __baseBuildVatReportSheet_v614 = typeof buildVatReportSheet_v611_ === 'function' ? buildVatReportSheet_v611_ : null;

aggregateSalesByBrand_v611_ = function() {
  var salesAgg = __baseAggregateSalesByBrand_v614 ? __baseAggregateSalesByBrand_v614.apply(this, arguments) : aggregateSalesNetByBrand_v613_();
  return enrichPurchaseAndProfit_v614_(salesAgg);
};

rebuildDashboardBrandBased_v611_ = function(salesAgg, filterAgg) {
  salesAgg = enrichPurchaseAndProfit_v614_(salesAgg || aggregateSalesByBrand_v611_());
  var result = __baseRebuildDashboardBrandBased_v614 ? __baseRebuildDashboardBrandBased_v614.call(this, salesAgg, filterAgg || aggregateFiltersByBrand_v611_()) : null;
  fixYearMonthDayFormats_v614_();
  return result;
};

buildBrandMarginSheet_v611_ = function(salesAgg) {
  salesAgg = enrichPurchaseAndProfit_v614_(salesAgg || aggregateSalesByBrand_v611_());
  var result = __baseBuildBrandMarginSheet_v614 ? __baseBuildBrandMarginSheet_v614.call(this, salesAgg) : buildBrandMarginSheetNetSales_v613_(salesAgg);
  fixYearMonthDayFormats_v614_();
  return result;
};

buildVatReportSheet_v611_ = function(salesAgg) {
  salesAgg = enrichPurchaseAndProfit_v614_(salesAgg || aggregateSalesByBrand_v611_());
  var result = __baseBuildVatReportSheet_v614 ? __baseBuildVatReportSheet_v614.call(this, salesAgg) : buildVatBreakdownSheetsNetSales_v613_(salesAgg);
  fixYearMonthDayFormats_v614_();
  return result;
};

function enrichPurchaseAndProfit_v614_(salesAgg) {
  salesAgg = salesAgg || {};
  if (salesAgg.__v614PurchaseEnriched) return salesAgg;
  salesAgg.__v614PurchaseEnriched = true;

  var rawIndex = buildRawPurchaseLookup_v614_();
  var purchaseByBrand = {};
  var purchaseByMonthBrand = {};
  var totalPurchase = 0;
  var foundPurchase = false;

  (salesAgg.detailRows || []).forEach(function(d) {
    var sourceRow = d.sourceRow || {};
    var purchase = getPurchaseAmountFromRow_v614_(sourceRow);
    if (!purchase) {
      purchase = lookupPurchaseAmount_v614_(rawIndex, d.orderNo, d.productNo, d.productName, d.netSales, d.qty);
    }
    purchase = toNumber_(purchase);
    d.purchase = purchase;
    d.netProfit = (d.settlement || 0) - purchase;
    if (purchase) foundPurchase = true;

    var brand = canonicalBrand_v613_(d.brand || '브랜드미확인');
    var key = normalizeBrandKey_v613_(brand);
    purchaseByBrand[key] = (purchaseByBrand[key] || 0) + purchase;
    var monthKey = String(d.year || '') + '-' + ('0' + String(d.month || '')).slice(-2);
    var mk = key + '||' + monthKey;
    purchaseByMonthBrand[mk] = (purchaseByMonthBrand[mk] || 0) + purchase;
    totalPurchase += purchase;
  });

  Object.keys(salesAgg.byBrand || {}).forEach(function(key) {
    var b = salesAgg.byBrand[key];
    var purchase = purchaseByBrand[key] || 0;
    b.purchaseAmount = purchase;
    b.netProfitAmount = (b.settlementAmount || 0) - purchase;
    b.netProfitRate = b.netSalesAmount ? b.netProfitAmount / b.netSalesAmount : 0;
    b.marginAmount = b.netProfitAmount;
    b.marginRate = b.netProfitRate;

    Object.keys(b.months || {}).forEach(function(month) {
      var m = b.months[month];
      var mp = purchaseByMonthBrand[key + '||' + month] || 0;
      m.purchase = mp;
      m.netProfitAmount = (m.settlement || 0) - mp;
      m.netProfitRate = m.sales ? m.netProfitAmount / m.sales : 0;
    });
  });

  salesAgg.totalPurchase = totalPurchase;
  salesAgg.totalNetProfit = (salesAgg.totalSettlement || 0) - totalPurchase;
  salesAgg.totalNetProfitRate = salesAgg.totalNetSales ? salesAgg.totalNetProfit / salesAgg.totalNetSales : 0;
  if (foundPurchase || rawIndex.hasPurchaseColumn) salesAgg.hasPurchaseColumn = true;

  try {
    log_('patch_purchase_v614', 'totalPurchase=' + totalPurchase + ' / hasPurchase=' + (salesAgg.hasPurchaseColumn ? 'Y' : 'N') + ' / rawPurchaseHeaders=' + rawIndex.purchaseHeaders.join(','));
  } catch (e) {}

  return salesAgg;
}

function buildRawPurchaseLookup_v614_() {
  var result = {
    byOrderProduct: {},
    byOrderProductName: {},
    byOrderSingle: {},
    orderCounts: {},
    hasPurchaseColumn: false,
    purchaseHeaders: []
  };

  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SALES_IN);
  if (!sheet || sheet.getLastRow() < 2) return result;

  var values = sheet.getDataRange().getValues();
  var header = values[0].map(function(h) { return String(h || '').trim(); });
  var purchaseCols = [];
  header.forEach(function(h, idx) {
    if (isPurchaseHeader_v614_(h)) {
      purchaseCols.push(idx);
      result.purchaseHeaders.push(h);
    }
  });
  result.hasPurchaseColumn = purchaseCols.length > 0;

  for (var r = 1; r < values.length; r++) {
    var obj = {};
    for (var c = 0; c < header.length; c++) if (header[c]) obj[header[c]] = values[r][c];
    var purchase = getPurchaseAmountFromRow_v614_(obj);
    if (!purchase) continue;

    var orderNo = String(getFirstValue_v613_(obj, ['마켓주문번호','주문번호','주문ID','주문번호_원본']) || '').trim();
    var productNo = String(getFirstValue_v613_(obj, ['사이트상품번호','마켓상품번호','상품번호','상품ID','상품코드']) || '').trim();
    var productName = String(getFirstValue_v613_(obj, ['마켓상품명','원문상품명','상품명','상품명_원본','옵션명']) || '').trim();
    if (!orderNo) continue;

    result.orderCounts[orderNo] = (result.orderCounts[orderNo] || 0) + 1;
    if (productNo) addLookupAmount_v614_(result.byOrderProduct, orderNo + '||' + productNo, purchase);
    if (productName) addLookupAmount_v614_(result.byOrderProductName, orderNo + '||' + normalizeTextKey_v614_(productName), purchase);
    addLookupAmount_v614_(result.byOrderSingle, orderNo, purchase);
  }

  return result;
}

function lookupPurchaseAmount_v614_(index, orderNo, productNo, productName, netSales, qty) {
  orderNo = String(orderNo || '').trim();
  productNo = String(productNo || '').trim();
  productName = String(productName || '').trim();
  if (!orderNo) return 0;
  if (productNo && index.byOrderProduct[orderNo + '||' + productNo]) return index.byOrderProduct[orderNo + '||' + productNo];
  if (productName && index.byOrderProductName[orderNo + '||' + normalizeTextKey_v614_(productName)]) return index.byOrderProductName[orderNo + '||' + normalizeTextKey_v614_(productName)];

  // 주문번호 하나에 원본 행이 하나뿐일 때만 주문번호 단독 fallback 사용
  if ((index.orderCounts[orderNo] || 0) === 1 && index.byOrderSingle[orderNo]) return index.byOrderSingle[orderNo];
  return 0;
}

function addLookupAmount_v614_(map, key, amount) {
  map[key] = (map[key] || 0) + toNumber_(amount);
}

function getPurchaseAmountFromRow_v614_(obj) {
  if (!obj) return 0;
  var exactNames = [
    '매입금액','총매입금액','매입금액합계','매입액','매입가','매입가합계','상품매입금액','상품매입가','총상품매입금액',
    '구매금액','총구매금액','구매금액합계','구매가','구매가합계','구입금액','발주금액',
    '원가','원가금액','상품원가','총원가','원가합계',
    '롯데온매입금액','롯데온_매입금액','LOTTEON매입금액',
    'purchase_amount','lotteon_purchase_amount','buy_amount','buy_price','cost','cost_amount','purchasePrice','purchaseAmount'
  ];
  for (var i = 0; i < exactNames.length; i++) {
    if (Object.prototype.hasOwnProperty.call(obj, exactNames[i])) {
      var v = toNumber_(obj[exactNames[i]]);
      if (v) return v;
    }
  }

  var best = 0;
  Object.keys(obj).forEach(function(k) {
    if (best) return;
    if (!isPurchaseHeader_v614_(k)) return;
    var v = toNumber_(obj[k]);
    if (v) best = v;
  });
  return best;
}

function isPurchaseHeader_v614_(header) {
  var raw = String(header || '').trim();
  if (!raw) return false;
  var h = raw.toLowerCase().replace(/\s+/g, '').replace(/_/g, '');

  // 구매자/고객/주소 등은 제외
  if (/구매자|주문자|수령|고객|buyer|customer|receiver|name|주소|전화|연락|수량|qty|count|율|rate|%/.test(h)) return false;
  // 매출/정산/수수료/부가세/배송비 등은 제외
  if (/매출|판매|결제|정산|수수료|부가세|vat|배송|택배|쿠폰|포인트|마켓|marketfee|settlement|sales|payment|fee|shipping/.test(h)) return false;

  if (/매입|구매금액|구매가|구입|발주금액|원가|purchase|buyamount|buyprice|cost/.test(h)) return true;
  return false;
}

function fixYearMonthDayFormats_v614_() {
  ['부가세_신고자료', '부가세_주문번호별'].forEach(function(sheetName) {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;
    var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    ['년','월','일'].forEach(function(name) {
      var col = header.indexOf(name) + 1;
      if (col > 0) sheet.getRange(2, col, sheet.getLastRow() - 1, 1).setNumberFormat('0');
    });
    var orderCol = header.indexOf('주문번호') + 1;
    if (orderCol > 0) sheet.getRange(2, orderCol, sheet.getLastRow() - 1, 1).setNumberFormat('@');
  });
}

function normalizeTextKey_v614_(v) {
  return String(v || '').toLowerCase().replace(/\s+/g, '').replace(/[\[\]\(\)\{\}\-_]/g, '').trim();
}
