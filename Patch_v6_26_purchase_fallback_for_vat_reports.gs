/**
 * LOTTEON v6.26 purchase fallback for VAT reports
 *
 * 문제:
 * - 부가세_상품별 시트에서 실제 매입금액이 있는 상품인데도 매입금액이 0으로 표시되는 사례가 있었습니다.
 * - 원인은 매출 집계 detailRows의 주문번호/상품번호 기준 매입 보강이 실패했을 때,
 *   상품번호+매출액 또는 상품명+매출액 기준 fallback이 없어서 매입금액을 못 가져오는 경우입니다.
 *
 * 수정:
 * - 매출데이터_붙여넣기 + 매출데이터_정리 양쪽을 모두 읽어 매입금액 lookup index를 만듭니다.
 * - 기존 exact match 유지:
 *   1) 주문번호 + 상품번호 + 순수매출액
 *   2) 주문번호 + 상품번호
 *   3) 주문번호 + 상품명 + 순수매출액
 *   4) 주문번호 + 상품명
 * - 추가 fallback:
 *   5) 상품번호 + 순수매출액
 *   6) 상품명 + 순수매출액
 *   7) 상품번호 단독: 해당 상품번호의 매입금액 후보가 1종류일 때만
 *   8) 상품명 단독: 해당 상품명의 매입금액 후보가 1종류일 때만
 * - fallback으로 매입금액을 찾으면 detailRows/브랜드 집계/부가세 집계의 매입금액과 이익을 재계산합니다.
 * - 아직 매입금액이 0인 행은 매입금액_미매칭 시트에 진단용으로 남깁니다.
 */

var LOTTEON_PATCH_V626_PURCHASE_FALLBACK_FOR_VAT_REPORTS_LOADED = true;

var __baseAggregateSalesByBrand_v626 = typeof aggregateSalesByBrand_v611_ === 'function' ? aggregateSalesByBrand_v611_ : null;
var __baseBuildVatBreakdownSheetsVatCredit_v626 = typeof buildVatBreakdownSheetsVatCredit_v621_ === 'function' ? buildVatBreakdownSheetsVatCredit_v621_ : null;
var __baseBuildBrandMarginSheet_v626 = typeof buildBrandMarginSheet_v611_ === 'function' ? buildBrandMarginSheet_v611_ : null;

aggregateSalesByBrand_v611_ = function() {
  var salesAgg = __baseAggregateSalesByBrand_v626 ? __baseAggregateSalesByBrand_v626.apply(this, arguments) : null;
  return enrichMissingPurchaseFallback_v626_(salesAgg);
};

buildVatBreakdownSheetsVatCredit_v621_ = function(salesAgg) {
  salesAgg = enrichMissingPurchaseFallback_v626_(salesAgg || aggregateSalesByBrand_v611_());
  return __baseBuildVatBreakdownSheetsVatCredit_v626 ? __baseBuildVatBreakdownSheetsVatCredit_v626.call(this, salesAgg) : null;
};

buildBrandMarginSheet_v611_ = function(salesAgg) {
  salesAgg = enrichMissingPurchaseFallback_v626_(salesAgg || aggregateSalesByBrand_v611_());
  return __baseBuildBrandMarginSheet_v626 ? __baseBuildBrandMarginSheet_v626.call(this, salesAgg) : null;
};

function enrichMissingPurchaseFallback_v626_(salesAgg) {
  salesAgg = salesAgg || {};
  if (salesAgg.__v626PurchaseFallbackDone) return salesAgg;
  salesAgg.__v626PurchaseFallbackDone = true;

  var details = salesAgg.detailRows || [];
  var missingBefore = 0;
  for (var i = 0; i < details.length; i++) {
    if (!toNumber_(details[i].purchase)) missingBefore++;
  }
  if (!missingBefore) return salesAgg;

  var index = buildPurchaseFallbackIndex_v626_();
  var matched = [];
  var unmatched = [];

  details.forEach(function(d, idx) {
    if (toNumber_(d.purchase)) return;
    var picked = pickPurchaseFallback_v626_(index, d);
    if (picked && picked.amount) {
      var amount = toNumber_(picked.amount);
      d.purchase = amount;
      d.purchaseMatchType = picked.type;
      d.purchaseMatchKey = picked.key;
      d.estimatedProfit = (toNumber_(d.settlementBasis || d.settlement) - amount);
      d.netProfit = d.estimatedProfit;
      matched.push([idx + 1, d.orderNo || '', d.productNo || '', d.brand || '', compactTextForDiag_v626_(d.productName || '', 80), d.netSales || 0, amount, picked.type, picked.key]);
    } else {
      unmatched.push([idx + 1, d.orderNo || '', d.productNo || '', d.brand || '', compactTextForDiag_v626_(d.productName || '', 80), d.netSales || 0, d.qty || '', '매입금액 후보 없음']);
    }
  });

  if (matched.length) recomputeSalesAggPurchaseProfit_v626_(salesAgg);
  writePurchaseFallbackDiagnostics_v626_(matched, unmatched, index);

  try {
    log_('purchase_fallback_v626', 'missingBefore=' + missingBefore + ' / matched=' + matched.length + ' / unmatched=' + unmatched.length + ' / sheets=' + index.sourceSheets.join(','));
  } catch (e) {}

  return salesAgg;
}

function buildPurchaseFallbackIndex_v626_() {
  var ss = SpreadsheetApp.getActive();
  var names = uniqueArray_v626_([
    CONFIG.SHEETS.SALES_IN,
    CONFIG.SHEETS.SALES_CLEAN,
    '매출데이터_붙여넣기',
    '매출데이터_정리'
  ]);
  var index = {
    orderProductSales: {},
    orderProduct: {},
    orderNameSales: {},
    orderName: {},
    productSales: {},
    nameSales: {},
    productOnly: {},
    nameOnly: {},
    sourceSheets: [],
    purchaseHeaders: []
  };

  names.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return;
    index.sourceSheets.push(name);
    var values = sheet.getDataRange().getValues();
    var header = values[0].map(function(h) { return String(h || '').trim(); });
    header.forEach(function(h) { if (isPurchaseHeader_v614_(h) && index.purchaseHeaders.indexOf(h) < 0) index.purchaseHeaders.push(h); });

    for (var r = 1; r < values.length; r++) {
      var obj = {};
      for (var c = 0; c < header.length; c++) if (header[c]) obj[header[c]] = values[r][c];
      var purchase = toNumber_(getPurchaseAmountFromRow_v614_(obj));
      if (!purchase) continue;

      var orderNo = stringKey_v626_(getFirstValue_v613_(obj, ['마켓주문번호','주문번호','주문ID','주문번호_원본','주문 번호','마켓 주문번호']));
      var productNo = stringKey_v626_(getFirstValue_v613_(obj, ['사이트상품번호','마켓상품번호','상품번호','상품ID','상품코드','상품 번호','사이트 상품번호']));
      var productName = stringKey_v626_(getFirstValue_v613_(obj, ['마켓상품명','원문상품명','상품명','상품명_원본','옵션명','상품 명']));
      var nameKey = normalizeTextKey_v614_(productName);
      var salesAmount = toNumber_(getFirstValue_v613_(obj, ['순수매출액','결제금액합계','결제금액','매출액','판매금액','총결제금액','상품금액']));
      var salesKey = moneyKey_v626_(salesAmount);

      if (orderNo && productNo && salesAmount) addPurchaseIndexValue_v626_(index.orderProductSales, orderNo + '||' + productNo + '||' + salesKey, purchase);
      if (orderNo && productNo) addPurchaseIndexValue_v626_(index.orderProduct, orderNo + '||' + productNo, purchase);
      if (orderNo && nameKey && salesAmount) addPurchaseIndexValue_v626_(index.orderNameSales, orderNo + '||' + nameKey + '||' + salesKey, purchase);
      if (orderNo && nameKey) addPurchaseIndexValue_v626_(index.orderName, orderNo + '||' + nameKey, purchase);
      if (productNo && salesAmount) addPurchaseIndexValue_v626_(index.productSales, productNo + '||' + salesKey, purchase);
      if (nameKey && salesAmount) addPurchaseIndexValue_v626_(index.nameSales, nameKey + '||' + salesKey, purchase);
      if (productNo) addPurchaseIndexValue_v626_(index.productOnly, productNo, purchase);
      if (nameKey) addPurchaseIndexValue_v626_(index.nameOnly, nameKey, purchase);
    }
  });
  return index;
}

function pickPurchaseFallback_v626_(index, d) {
  var orderNo = stringKey_v626_(d.orderNo);
  var productNo = stringKey_v626_(d.productNo);
  var nameKey = normalizeTextKey_v614_(d.productName || '');
  var sales = toNumber_(d.netSales || 0);
  var salesKey = moneyKey_v626_(sales);
  var candidates = [];

  if (orderNo && productNo && sales) candidates.push(['주문+상품번호+매출액', index.orderProductSales, orderNo + '||' + productNo + '||' + salesKey, false]);
  if (orderNo && productNo) candidates.push(['주문+상품번호', index.orderProduct, orderNo + '||' + productNo, false]);
  if (orderNo && nameKey && sales) candidates.push(['주문+상품명+매출액', index.orderNameSales, orderNo + '||' + nameKey + '||' + salesKey, false]);
  if (orderNo && nameKey) candidates.push(['주문+상품명', index.orderName, orderNo + '||' + nameKey, false]);
  if (productNo && sales) candidates.push(['상품번호+매출액', index.productSales, productNo + '||' + salesKey, false]);
  if (nameKey && sales) candidates.push(['상품명+매출액', index.nameSales, nameKey + '||' + salesKey, false]);
  if (productNo) candidates.push(['상품번호단독_단일후보', index.productOnly, productNo, true]);
  if (nameKey) candidates.push(['상품명단독_단일후보', index.nameOnly, nameKey, true]);

  for (var i = 0; i < candidates.length; i++) {
    var type = candidates[i][0];
    var map = candidates[i][1];
    var key = candidates[i][2];
    var requireUnique = candidates[i][3];
    var amount = getPurchaseIndexAmount_v626_(map, key, requireUnique);
    if (amount) return { amount: amount, type: type, key: key };
  }
  return null;
}

function addPurchaseIndexValue_v626_(map, key, amount) {
  if (!key) return;
  amount = toNumber_(amount);
  if (!amount) return;
  if (!map[key]) map[key] = { amount: 0, count: 0, distinct: {} };
  map[key].amount += amount;
  map[key].count += 1;
  map[key].distinct[String(Math.round(amount))] = true;
}

function getPurchaseIndexAmount_v626_(map, key, requireUnique) {
  var item = map && map[key];
  if (!item) return 0;
  if (requireUnique && Object.keys(item.distinct || {}).length > 1) return 0;
  var count = Math.max(1, toNumber_(item.count));
  return toNumber_(item.amount) / count;
}

function recomputeSalesAggPurchaseProfit_v626_(salesAgg) {
  var totalPurchase = 0;
  var totalProfit = 0;
  Object.keys(salesAgg.byBrand || {}).forEach(function(key) {
    var b = salesAgg.byBrand[key];
    b.purchaseAmount = 0;
    b.estimatedProfitAmount = 0;
    b.netProfitAmount = 0;
    b.marginAmount = 0;
    Object.keys(b.months || {}).forEach(function(mk) {
      b.months[mk].purchase = 0;
      b.months[mk].estimatedProfitAmount = 0;
      b.months[mk].netProfitAmount = 0;
    });
  });

  (salesAgg.detailRows || []).forEach(function(d) {
    var purchase = toNumber_(d.purchase);
    var settlementBasis = toNumber_(d.settlementBasis || d.settlement);
    var profit = settlementBasis - purchase;
    d.estimatedProfit = profit;
    d.netProfit = profit;
    totalPurchase += purchase;
    totalProfit += profit;

    var brandKey = normalizeBrandKey_v613_(canonicalBrand_v613_(d.brand || '브랜드미확인'));
    var b = salesAgg.byBrand && salesAgg.byBrand[brandKey];
    if (b) {
      b.purchaseAmount += purchase;
      b.estimatedProfitAmount += profit;
      b.netProfitAmount += profit;
      b.marginAmount += profit;
      b.estimatedProfitRate = b.netSalesAmount ? b.estimatedProfitAmount / b.netSalesAmount : 0;
      b.netProfitRate = b.estimatedProfitRate;
      b.marginRate = b.estimatedProfitRate;
      var monthKey = d.year && d.month ? d.year + '-' + ('0' + d.month).slice(-2) : '';
      if (b.months && b.months[monthKey]) {
        b.months[monthKey].purchase += purchase;
        b.months[monthKey].estimatedProfitAmount += profit;
        b.months[monthKey].netProfitAmount = b.months[monthKey].estimatedProfitAmount;
        b.months[monthKey].netProfitRate = b.months[monthKey].sales ? b.months[monthKey].estimatedProfitAmount / b.months[monthKey].sales : 0;
      }
    }
  });

  salesAgg.totalPurchase = totalPurchase;
  salesAgg.totalEstimatedProfit = totalProfit;
  salesAgg.totalNetProfit = totalProfit;
  salesAgg.totalEstimatedProfitRate = salesAgg.totalNetSales ? totalProfit / salesAgg.totalNetSales : 0;
  salesAgg.totalNetProfitRate = salesAgg.totalEstimatedProfitRate;
  salesAgg.hasPurchaseColumn = true;
}

function writePurchaseFallbackDiagnostics_v626_(matched, unmatched, index) {
  var ss = SpreadsheetApp.getActive();
  try {
    var missSheet = ss.getSheetByName('매입금액_미매칭') || ss.insertSheet('매입금액_미매칭');
    var missHeaders = ['행번호','주문번호','상품번호','브랜드명','상품명','순수매출액','판매수량','사유'];
    replaceSheetValues_v611_(missSheet, missHeaders, unmatched.slice(0, 500));
    formatMoneyColumns_v611_(missSheet, [6]);
    formatNumberColumns_v611_(missSheet, [1,7]);
  } catch (e) {}

  try {
    var logSheet = ss.getSheetByName('매입금액_보강로그') || ss.insertSheet('매입금액_보강로그');
    var logHeaders = ['행번호','주문번호','상품번호','브랜드명','상품명','순수매출액','보강매입금액','매칭방식','매칭키'];
    var rows = matched.slice(0, 500);
    rows.unshift(['요약','보강성공=' + matched.length,'미매칭=' + unmatched.length,'검색시트=' + index.sourceSheets.join(', '),'매입헤더=' + index.purchaseHeaders.join(', '),'','','','']);
    replaceSheetValues_v611_(logSheet, logHeaders, rows);
    formatMoneyColumns_v611_(logSheet, [6,7]);
  } catch (e) {}
}

function uniqueArray_v626_(arr) {
  var seen = {};
  var out = [];
  (arr || []).forEach(function(v) { v = String(v || '').trim(); if (v && !seen[v]) { seen[v] = true; out.push(v); } });
  return out;
}

function stringKey_v626_(v) {
  return String(v == null ? '' : v).trim();
}

function moneyKey_v626_(v) {
  return String(Math.round(toNumber_(v)));
}

function compactTextForDiag_v626_(text, limit) {
  text = String(text == null ? '' : text).replace(/\s+/g, ' ').trim();
  if (!limit || text.length <= limit) return text;
  return text.slice(0, limit - 1) + '…';
}
