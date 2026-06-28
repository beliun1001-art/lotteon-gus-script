/**
 * LOTTEON v6.29 customer/address grouping + unsettled restore patch
 *
 * 문제:
 * - v6.28 단일 원천 집계 전환 후 detailRows의 고객명/주소/쿠팡계정ID가 비어 있어
 *   부가세_고객별이 고객명미확인 1행으로만 집계되었습니다.
 * - 미정산_쿠팡계정별 시트도 detailRows에 쿠팡계정ID/고객명 정보가 없어 내용이 사라졌습니다.
 *
 * 수정:
 * - 매출데이터_붙여넣기 원본 행에서 고객명, 배송주소, 쿠팡계정ID를 추출합니다.
 * - 고객별 집계는 고객명 + 주소 기준으로 구분합니다.
 * - 주소 원문은 시트에 표시하지 않습니다.
 * - 동일 고객명에 주소가 여러 개면 주소구분을 주소1, 주소2처럼 번호로 표시합니다.
 * - 미정산_쿠팡계정별 시트를 v6.28 단일 원천 detailRows 기준으로 재생성합니다.
 */

var LOTTEON_PATCH_V629_CUSTOMER_ADDRESS_UNSETTLED_RESTORE_LOADED = true;
var LOTTEON_V629_SETTLEMENT_RATE = 0.901;

var __baseMakeSourceRow_v629 = typeof makeSourceRow_v628_ === 'function' ? makeSourceRow_v628_ : null;
var __baseSourceRowToDetail_v629 = typeof sourceRowToDetail_v628_ === 'function' ? sourceRowToDetail_v628_ : null;
var __baseBuildSingleSourceSalesAgg_v629 = typeof buildSingleSourceSalesAgg_v628_ === 'function' ? buildSingleSourceSalesAgg_v628_ : null;

makeSourceRow_v628_ = function(sheetRow, obj, sales, purchase) {
  var row = __baseMakeSourceRow_v629 ? __baseMakeSourceRow_v629.apply(this, arguments) : {};
  row.accountId = normalizeId_v629_(getAny_v628_(obj, [
    '원본계정ID','분석계정ID','쿠팡계정ID','계정ID','계정번호','판매자ID','판매자 아이디','쿠팡 ID','쿠팡ID'
  ]));
  if (!row.accountId) row.accountId = normalizeId_v629_(getAny_v628_(obj, ['계정','마켓계정','마켓 계정']));

  row.customerName = normalizeCustomerName_v629_(getAny_v628_(obj, [
    '수령인','수취인','받는분','받는사람','수령자','수취인명','배송수령인','배송 수령인',
    '주문자','주문자명','구매자','구매자명','고객명','고객 이름','수령인명'
  ]));
  row.address = normalizeAddress_v629_(getAny_v628_(obj, [
    '배송주소','수령인주소','수취인주소','주소','배송지주소','배송지 주소','수령주소',
    '받는분주소','받는사람주소','기본주소','상세주소','수취주소'
  ]));
  row.addressKey = normalizeAddressKey_v629_(row.address);
  row.customerAddressKey = (row.customerName || '고객명미확인') + '||' + (row.addressKey || '주소미확인');
  return row;
};

sourceRowToDetail_v628_ = function(row) {
  var d = __baseSourceRowToDetail_v629 ? __baseSourceRowToDetail_v629.apply(this, arguments) : {};
  d.accountId = row.accountId || '';
  d.customer = row.customerName || '고객명미확인';
  d.customerName = d.customer;
  d.addressKey = row.addressKey || '주소미확인';
  d.customerAddressKey = row.customerAddressKey || (d.customer + '||' + d.addressKey);
  d.addressGroup = '';
  d.customerDisplay = d.customer;
  return d;
};

buildSingleSourceSalesAgg_v628_ = function() {
  var agg = __baseBuildSingleSourceSalesAgg_v629 ? __baseBuildSingleSourceSalesAgg_v629.apply(this, arguments) : null;
  return applyCustomerAddressGroups_v629_(agg);
};

// 혹시 다른 patch가 aggregateSalesByBrand_v611_를 다시 감싸도 최종적으로 고객/주소 보강이 되도록 한 번 더 보정합니다.
var __baseAggregateSalesByBrand_v629 = typeof aggregateSalesByBrand_v611_ === 'function' ? aggregateSalesByBrand_v611_ : null;
aggregateSalesByBrand_v611_ = function() {
  var agg = __baseAggregateSalesByBrand_v629 ? __baseAggregateSalesByBrand_v629.apply(this, arguments) : buildSingleSourceSalesAgg_v628_();
  return applyCustomerAddressGroups_v629_(agg);
};

buildVatCustomerSingleSource_v628_ = function(salesAgg) {
  salesAgg = applyCustomerAddressGroups_v629_(salesAgg || buildSingleSourceSalesAgg_v628_());
  var map = groupDetails_v628_(salesAgg.detailRows, function(d) {
    return d.customerAddressKey || ((d.customer || '고객명미확인') + '||' + (d.addressKey || '주소미확인'));
  }, function(d) {
    return {
      customer: d.customer || '고객명미확인',
      addressGroup: d.addressGroup || '',
      customerDisplay: d.customerDisplay || d.customer || '고객명미확인',
      orderSet: {},
      productSet: {}
    };
  });

  var rows = Object.keys(map).map(function(k) {
    var g = map[k];
    var ss = splitVat_v628_(g.netSales);
    var ps = splitVat_v628_(g.purchase);
    var payableVat = ss.vat - ps.vat;
    return [
      g.customer,
      g.addressGroup || '',
      Object.keys(g.orderSet || {}).length,
      Object.keys(g.productSet || {}).length,
      g.qty,
      g.netSales,
      ss.supply,
      ss.vat,
      g.purchase,
      ps.supply,
      ps.vat,
      payableVat,
      g.profit,
      g.profit - payableVat
    ];
  }).sort(function(a, b) { return b[5] - a[5]; });

  var sheet = SpreadsheetApp.getActive().getSheetByName('부가세_고객별') || SpreadsheetApp.getActive().insertSheet('부가세_고객별');
  replaceSheetValuesSafe_v628_(sheet, [
    '고객명','주소구분','주문건수','상품수','판매수량','순수매출액','매출공급가액','매출부가세',
    '매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익'
  ], rows);
  formatOutputSheetBasic_v628_(sheet);
  return { rows: rows.length };
};

// v6.28 부가세 전체 생성 경로에서 고객별 시트가 반드시 v6.29 고객/주소 기준을 쓰도록 재정의합니다.
buildVatReportsSingleSource_v628_ = function(salesAgg) {
  salesAgg = applyCustomerAddressGroups_v629_(salesAgg || buildSingleSourceSalesAgg_v628_());
  buildVatDetailSingleSource_v628_(salesAgg);
  buildVatProductSingleSource_v628_(salesAgg);
  buildVatCustomerSingleSource_v628_(salesAgg);
  buildVatOrderSingleSource_v628_(salesAgg);
  buildUnsettledSettlementByAccountSheet_v629_(salesAgg);
  buildCustomerAddressGroupingDiagnostics_v629_(salesAgg);
  buildFinancialValidationSheets_v628_(salesAgg);
  return { rows: salesAgg.detailRows.length };
};

buildUnsettledSettlementByAccountSheet_v616_ = function(salesAgg) {
  return buildUnsettledSettlementByAccountSheet_v629_(salesAgg || buildSingleSourceSalesAgg_v628_());
};

function buildUnsettledSettlementByAccountSheet_v629_(salesAgg) {
  salesAgg = applyCustomerAddressGroups_v629_(salesAgg || buildSingleSourceSalesAgg_v628_());
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('미정산_쿠팡계정별') || ss.insertSheet('미정산_쿠팡계정별');
  var today = new Date();
  var rows = [];

  (salesAgg.detailRows || []).forEach(function(d) {
    var actual = num_v628_(d.actualSettlement);
    if (actual) return;
    var orderDate = d.orderDate || null;
    var elapsed = orderDate ? Math.floor((today.getTime() - orderDate.getTime()) / (1000 * 60 * 60 * 24)) : '';
    var over30 = elapsed !== '' && elapsed > 30;
    rows.push([
      d.accountId || '계정미확인',
      over30 ? '30일초과' : '미정산',
      1,
      over30 ? 1 : 0,
      d.dateText || '',
      elapsed,
      over30 ? d.orderNo : '',
      d.brand || '',
      d.customer || '고객명미확인',
      d.addressGroup || '',
      d.productNo || '',
      compactText_v628_(d.productName || '', 80),
      d.netSales || 0,
      d.estimatedSettlement || Math.round((d.netSales || 0) * LOTTEON_V629_SETTLEMENT_RATE),
      d.purchase || 0,
      d.estimatedProfit || 0,
      d.overdueStatus || d.unsettledStatus || '미정산'
    ]);
  });

  rows.sort(function(a, b) {
    return String(a[0]).localeCompare(String(b[0])) || String(b[4]).localeCompare(String(a[4])) || String(a[6]).localeCompare(String(b[6]));
  });

  replaceSheetValuesSafe_v628_(sheet, [
    '쿠팡계정ID','구분','미정산건수','30일초과건수','주문일','경과일','30일초과주문번호',
    '브랜드명','고객명','주소구분','상품번호','상품명','순수매출액','예상정산금액','매입금액','예상이익','비고'
  ], rows);
  formatOutputSheetBasic_v628_(sheet);
  return { rows: rows.length };
}

function applyCustomerAddressGroups_v629_(salesAgg) {
  salesAgg = salesAgg || {};
  if (salesAgg.__v629CustomerAddressApplied) return salesAgg;
  salesAgg.__v629CustomerAddressApplied = true;

  var details = salesAgg.detailRows || [];
  var byName = {};
  details.forEach(function(d) {
    d.customer = normalizeCustomerName_v629_(d.customer || d.customerName || '') || '고객명미확인';
    d.addressKey = normalizeAddressKey_v629_(d.addressKey || '') || '주소미확인';
    d.customerAddressKey = d.customer + '||' + d.addressKey;
    if (!byName[d.customer]) byName[d.customer] = {};
    byName[d.customer][d.addressKey] = true;
  });

  var addressNoByKey = {};
  Object.keys(byName).sort().forEach(function(name) {
    var keys = Object.keys(byName[name]).sort();
    keys.forEach(function(addrKey, idx) {
      addressNoByKey[name + '||' + addrKey] = keys.length > 1 ? ('주소' + (idx + 1)) : '';
    });
  });

  details.forEach(function(d) {
    d.addressGroup = addressNoByKey[d.customerAddressKey] || '';
    d.customerDisplay = d.addressGroup ? (d.customer + ' (' + d.addressGroup + ')') : d.customer;
  });

  salesAgg.customerAddressGroupCount = Object.keys(addressNoByKey).length;
  salesAgg.sameNameMultiAddressCount = Object.keys(byName).filter(function(name) { return Object.keys(byName[name]).length > 1; }).length;
  return salesAgg;
}

function buildCustomerAddressGroupingDiagnostics_v629_(salesAgg) {
  salesAgg = applyCustomerAddressGroups_v629_(salesAgg || buildSingleSourceSalesAgg_v628_());
  var byName = {};
  (salesAgg.detailRows || []).forEach(function(d) {
    var name = d.customer || '고객명미확인';
    if (!byName[name]) byName[name] = { addressGroups: {}, orders: {}, sales: 0 };
    byName[name].addressGroups[d.addressGroup || '단일주소'] = true;
    if (d.orderNo) byName[name].orders[d.orderNo] = true;
    byName[name].sales += d.netSales || 0;
  });
  var rows = Object.keys(byName).map(function(name) {
    var g = byName[name];
    return [name, Object.keys(g.addressGroups).length, Object.keys(g.orders).length, g.sales, Object.keys(g.addressGroups).join(', ')];
  }).sort(function(a, b) { return b[3] - a[3]; });

  var sheet = SpreadsheetApp.getActive().getSheetByName('고객주소_구분검증') || SpreadsheetApp.getActive().insertSheet('고객주소_구분검증');
  replaceSheetValuesSafe_v628_(sheet, ['고객명','주소구분수','주문건수','순수매출액','주소구분목록'], rows);
  formatOutputSheetBasic_v628_(sheet);
}

function normalizeCustomerName_v629_(v) {
  var s = String(v == null ? '' : v).trim();
  s = s.replace(/\s+/g, ' ');
  return s || '';
}

function normalizeAddress_v629_(v) {
  var s = String(v == null ? '' : v).trim();
  s = s.replace(/\s+/g, ' ');
  return s;
}

function normalizeAddressKey_v629_(v) {
  var s = normalizeAddress_v629_(v);
  if (!s) return '';
  return s.replace(/[\s,.-]/g, '').toLowerCase();
}

function normalizeId_v629_(v) {
  return String(v == null ? '' : v).replace(/,/g, '').trim();
}
