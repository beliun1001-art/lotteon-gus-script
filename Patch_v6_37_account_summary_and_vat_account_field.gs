/**
 * LOTTEON v6.37 account summary + VAT account field
 *
 * 확정 기준:
 * - 사업자별(계정별) 써머리 시트를 새로 생성합니다.
 * - 시트명: 사업자별_계정별_써머리
 * - 집계 기준은 v6.28+ 단일 원천 detailRows입니다.
 * - 부가세_신고자료 시트에는 쿠팡계정ID 필드를 열로 추가합니다.
 * - 계정이 비어 있으면 원본 매출데이터_붙여넣기 sourceRow 기준으로 복구합니다.
 *
 * 사업자별_계정별_써머리 주요 컬럼:
 * - 쿠팡계정ID, 주문건수, 고객수, 판매수량, 매출상품수
 * - 총매출액, 취소/반품매출, 순수매출액
 * - 실제정산금액, 예상정산금액(미정산), 정산기준금액
 * - 마켓수수료/비용, 마켓수수료율
 * - 매입금액, 예상이익, 예상이익률
 * - 매출부가세, 매입부가세, 납부예상부가세
 * - 부가세반영예상이익, 부가세반영이익률
 * - 미정산주문수, 30일초과미정산건수, 비고
 */

var LOTTEON_PATCH_V637_ACCOUNT_SUMMARY_AND_VAT_ACCOUNT_FIELD_LOADED = true;

var __baseBuildVatReportsSingleSource_v637 = typeof buildVatReportsSingleSource_v628_ === 'function' ? buildVatReportsSingleSource_v628_ : null;
var __baseBuildVatDetailSingleSource_v637 = typeof buildVatDetailSingleSource_v628_ === 'function' ? buildVatDetailSingleSource_v628_ : null;
var __baseGenerateVatReportsFullSeparated_v637 = typeof generateVatReportsFullSeparated_v622 === 'function' ? generateVatReportsFullSeparated_v622 : null;
var __baseRefreshDashboardFastOnly_v637 = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;
var __baseApplyDisplayStandardsOnlyFast_v637 = typeof applyDisplayStandardsOnlyFast_v623 === 'function' ? applyDisplayStandardsOnlyFast_v623 : null;
var __baseApplyFastOutputSheetFormatting_v637 = typeof applyFastOutputSheetFormatting_v620_ === 'function' ? applyFastOutputSheetFormatting_v620_ : null;

buildVatReportsSingleSource_v628_ = function(salesAgg) {
  salesAgg = prepareAccountSummarySalesAgg_v637_(salesAgg || buildSingleSourceSalesAgg_v628_());
  buildVatDetailSingleSource_v628_(salesAgg);
  buildVatProductSingleSource_v628_(salesAgg);
  buildVatCustomerSingleSource_v628_(salesAgg);
  buildVatOrderSingleSource_v628_(salesAgg);
  if (typeof buildUnsettledSettlementByAccountSheet_v629_ === 'function') buildUnsettledSettlementByAccountSheet_v629_(salesAgg);
  if (typeof buildCustomerAddressGroupingDiagnostics_v629_ === 'function') buildCustomerAddressGroupingDiagnostics_v629_(salesAgg);
  buildAccountSummarySheet_v637_(salesAgg);
  if (typeof buildFinancialValidationSheets_v628_ === 'function') buildFinancialValidationSheets_v628_(salesAgg);
  return { rows: (salesAgg.detailRows || []).length };
};

buildVatDetailSingleSource_v628_ = function(salesAgg) {
  salesAgg = prepareAccountSummarySalesAgg_v637_(salesAgg || buildSingleSourceSalesAgg_v628_());
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('부가세_신고자료') || ss.insertSheet('부가세_신고자료');
  var headers = [
    '날짜','쿠팡계정ID','주문번호','고객명','브랜드명','상품번호','상품명','판매수량',
    '순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료/비용',
    '매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','비고'
  ];
  var rows = (salesAgg.detailRows || []).map(function(d) { return vatDetailRowWithAccount_v637_(d); });
  replaceSheetValuesSafe_v637_(sheet, headers, rows);
  formatAccountOutputSheet_v637_(sheet);
  return { rows: rows.length };
};

generateVatReportsFullSeparated_v622 = function() {
  var result = __baseGenerateVatReportsFullSeparated_v637 ? __baseGenerateVatReportsFullSeparated_v637.apply(this, arguments) : null;
  try {
    var salesAgg = prepareAccountSummarySalesAgg_v637_(buildSingleSourceSalesAgg_v628_());
    buildAccountSummarySheet_v637_(salesAgg);
    ensureVatDetailHasAccountField_v637_(salesAgg);
  } catch (e) { try { log_('patch_v637_after_vat_error', String(e && e.message ? e.message : e)); } catch(ignore) {} }
  return result || { ok: true };
};

refreshDashboardFastOnly = function() {
  var result = __baseRefreshDashboardFastOnly_v637 ? __baseRefreshDashboardFastOnly_v637.apply(this, arguments) : null;
  try { buildAccountSummarySheet_v637_(prepareAccountSummarySalesAgg_v637_(buildSingleSourceSalesAgg_v628_())); } catch (e) {}
  return result || { ok: true };
};

applyDisplayStandardsOnlyFast_v623 = function() {
  var result = __baseApplyDisplayStandardsOnlyFast_v637 ? __baseApplyDisplayStandardsOnlyFast_v637.apply(this, arguments) : null;
  try {
    var salesAgg = prepareAccountSummarySalesAgg_v637_(buildSingleSourceSalesAgg_v628_());
    buildAccountSummarySheet_v637_(salesAgg);
    ensureVatDetailHasAccountField_v637_(salesAgg);
    SpreadsheetApp.getUi().alert('표시서식 정리 완료\n\n사업자별_계정별_써머리와 부가세_신고자료 쿠팡계정ID 열을 확인했습니다.');
  } catch (e) {}
  return result || { ok: true };
};

applyFastOutputSheetFormatting_v620_ = function() {
  var result = __baseApplyFastOutputSheetFormatting_v637 ? __baseApplyFastOutputSheetFormatting_v637.apply(this, arguments) : null;
  try { formatAccountOutputSheet_v637_(SpreadsheetApp.getActive().getSheetByName('사업자별_계정별_써머리')); } catch (e) {}
  try { formatAccountOutputSheet_v637_(SpreadsheetApp.getActive().getSheetByName('부가세_신고자료')); } catch (e) {}
  return result || { ok: true };
};

function prepareAccountSummarySalesAgg_v637_(salesAgg) {
  salesAgg = salesAgg || {};
  if (salesAgg.__v637AccountPrepared) return salesAgg;
  salesAgg.__v637AccountPrepared = true;
  var accountMap = buildAccountBySourceRowForSummary_v637_();
  (salesAgg.detailRows || []).forEach(function(d) {
    var account = normalizeAccount_v637_(d.accountId);
    if (!isValidAccount_v637_(account)) account = accountMap[d.sourceRow] || account;
    d.accountId = account || '계정미확인';
  });
  return salesAgg;
}

function vatDetailRowWithAccount_v637_(d) {
  var salesSplit = splitVatSafe_v637_(d.netSales);
  var purchaseSplit = splitVatSafe_v637_(d.purchase);
  var payableVat = salesSplit.vat - purchaseSplit.vat;
  return [
    d.dateText || '',
    d.accountId || '계정미확인',
    d.orderNo || '',
    d.customer || '',
    d.brand || '',
    d.productNo || '',
    d.productName || '',
    num_v637_(d.qty),
    num_v637_(d.netSales),
    salesSplit.supply,
    salesSplit.vat,
    num_v637_(d.settlementBasis),
    num_v637_(d.marketFee),
    num_v637_(d.purchase),
    purchaseSplit.supply,
    purchaseSplit.vat,
    payableVat,
    num_v637_(d.estimatedProfit),
    num_v637_(d.estimatedProfit) - payableVat,
    d.note || 'v6.37 계정필드 포함'
  ];
}

function buildAccountSummarySheet_v637_(salesAgg) {
  salesAgg = prepareAccountSummarySalesAgg_v637_(salesAgg || buildSingleSourceSalesAgg_v628_());
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('사업자별_계정별_써머리') || ss.insertSheet('사업자별_계정별_써머리');
  var map = {};

  (salesAgg.detailRows || []).forEach(function(d) {
    var account = d.accountId || '계정미확인';
    if (!map[account]) map[account] = initAccountSummary_v637_(account);
    addDetailToAccountSummary_v637_(map[account], d);
  });

  var rows = Object.keys(map).sort().map(function(account) {
    var a = finalizeAccountSummary_v637_(map[account]);
    return [
      a.accountId,
      a.orderCount,
      a.customerCount,
      a.quantity,
      a.productCount,
      a.grossSalesAmount,
      a.cancelSalesAmount,
      a.netSalesAmount,
      a.actualSettlementAmount,
      a.estimatedSettlementAmount,
      a.settlementBasisAmount,
      a.marketFeeAmount,
      a.marketFeeRate,
      a.purchaseAmount,
      a.estimatedProfitAmount,
      a.estimatedProfitRate,
      a.salesVat,
      a.purchaseVat,
      a.payableVat,
      a.afterVatProfit,
      a.afterVatProfitRate,
      a.unsettledOrderCount,
      a.overdueUnsettledOrderCount,
      'v6.37 계정별 단일 원천'
    ];
  });

  var total = buildAccountSummaryTotalRow_v637_(rows);
  if (rows.length) rows.push(total);

  replaceSheetValuesSafe_v637_(sheet, [
    '쿠팡계정ID','주문건수','고객수','판매수량','매출상품수','총매출액','취소/반품매출','순수매출액',
    '실제정산금액','예상정산금액(미정산)','정산기준금액','마켓수수료/비용','마켓수수료율',
    '매입금액','예상이익','예상이익률','매출부가세','매입부가세','납부예상부가세',
    '부가세반영예상이익','부가세반영이익률','미정산주문수','30일초과미정산건수','비고'
  ], rows);
  formatAccountOutputSheet_v637_(sheet);
  try { log_('patch_v637_account_summary', 'rows=' + rows.length); } catch (e) {}
  return { rows: rows.length };
}

function initAccountSummary_v637_(account) {
  return {
    accountId: account,
    orderSet: {}, customerSet: {}, productSet: {}, unsettledOrderSet: {}, overdueUnsettledOrderSet: {},
    quantity: 0,
    grossSalesAmount: 0,
    cancelSalesAmount: 0,
    netSalesAmount: 0,
    actualSettlementAmount: 0,
    estimatedSettlementAmount: 0,
    settlementBasisAmount: 0,
    marketFeeAmount: 0,
    purchaseAmount: 0,
    estimatedProfitAmount: 0
  };
}

function addDetailToAccountSummary_v637_(a, d) {
  if (d.orderNo) a.orderSet[d.orderNo] = true;
  if (d.customerAddressKey) a.customerSet[d.customerAddressKey] = true;
  else if (d.customer) a.customerSet[d.customer] = true;
  var productKey = (d.productNo || '') + '|' + (d.productName || '');
  if (productKey !== '|') a.productSet[productKey] = true;
  if (d.unsettledStatus && d.orderNo) a.unsettledOrderSet[d.orderNo] = true;
  if (d.overdueStatus && d.orderNo) a.overdueUnsettledOrderSet[d.orderNo] = true;
  a.quantity += num_v637_(d.qty);
  a.grossSalesAmount += num_v637_(d.netSales);
  a.netSalesAmount += num_v637_(d.netSales);
  a.actualSettlementAmount += num_v637_(d.actualSettlement);
  a.estimatedSettlementAmount += num_v637_(d.estimatedSettlement);
  a.settlementBasisAmount += num_v637_(d.settlementBasis);
  a.marketFeeAmount += num_v637_(d.marketFee);
  a.purchaseAmount += num_v637_(d.purchase);
  a.estimatedProfitAmount += num_v637_(d.estimatedProfit);
}

function finalizeAccountSummary_v637_(a) {
  a.orderCount = Object.keys(a.orderSet).length;
  a.customerCount = Object.keys(a.customerSet).length;
  a.productCount = Object.keys(a.productSet).length;
  a.unsettledOrderCount = Object.keys(a.unsettledOrderSet).length;
  a.overdueUnsettledOrderCount = Object.keys(a.overdueUnsettledOrderSet).length;
  a.marketFeeRate = a.netSalesAmount ? a.marketFeeAmount / a.netSalesAmount : 0;
  a.estimatedProfitRate = a.netSalesAmount ? a.estimatedProfitAmount / a.netSalesAmount : 0;
  var salesSplit = splitVatSafe_v637_(a.netSalesAmount);
  var purchaseSplit = splitVatSafe_v637_(a.purchaseAmount);
  a.salesVat = salesSplit.vat;
  a.purchaseVat = purchaseSplit.vat;
  a.payableVat = a.salesVat - a.purchaseVat;
  a.afterVatProfit = a.estimatedProfitAmount - a.payableVat;
  a.afterVatProfitRate = a.netSalesAmount ? a.afterVatProfit / a.netSalesAmount : 0;
  return a;
}

function buildAccountSummaryTotalRow_v637_(rows) {
  var t = new Array(24).fill('');
  t[0] = '합계';
  [1,2,3,4,5,6,7,8,9,10,11,13,14,16,17,18,19,21,22].forEach(function(idx) {
    t[idx] = rows.reduce(function(sum, row) { return sum + num_v637_(row[idx]); }, 0);
  });
  t[12] = t[7] ? t[11] / t[7] : 0;
  t[15] = t[7] ? t[14] / t[7] : 0;
  t[20] = t[7] ? t[19] / t[7] : 0;
  t[23] = '계정별 합계';
  return t;
}

function ensureVatDetailHasAccountField_v637_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('부가세_신고자료');
  if (!sheet || sheet.getLastRow() < 1) return { skipped: true };
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h || '').replace(/\s+/g, '').trim(); });
  if (header.indexOf('쿠팡계정ID') >= 0) return { ok: true, already: true };
  // 계정열이 없으면 안전하게 전체 재생성합니다.
  return buildVatDetailSingleSource_v628_(salesAgg || buildSingleSourceSalesAgg_v628_());
}

function buildAccountBySourceRowForSummary_v637_() {
  if (typeof sourceRowAccountMap_v636_ === 'function') return sourceRowAccountMap_v636_();
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SALES_IN) || ss.getSheetByName('매출데이터_붙여넣기');
  var out = {};
  if (!sheet || sheet.getLastRow() < 2) return out;
  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function(h) { return String(h || '').trim(); });
  for (var r = 1; r < values.length; r++) {
    var account = detectAccountForSummary_v637_(headers, values[r]);
    if (account) out[r + 1] = account;
  }
  return out;
}

function detectAccountForSummary_v637_(headers, row) {
  var textAll = row.map(function(v){ return String(v == null ? '' : v); }).join(' ');
  var direct = textAll.match(/beliun\d{4}/i);
  if (direct) return direct[0];
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    var v = normalizeAccount_v637_(row[i]);
    if (/계정|account|아이디|ID|판매자/i.test(h)) {
      var mapped = mapAccountNo_v637_(v);
      if (mapped) return mapped;
    }
  }
  var filter = textAll.match(/롯백[_\s-]*(0?1|0?2|0?3|0?4)/);
  if (filter) {
    var code = String(filter[1]).replace(/^0/, '');
    if (code === '1') return 'beliun1021';
    if (code === '2') return 'beliun1024';
    if (code === '3') return 'beliun1023';
    if (code === '4') return 'beliun1024';
  }
  return '';
}

function formatAccountOutputSheet_v637_(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || '').replace(/\n/g, '').replace(/\s+/g, '').trim(); });
  try {
    sheet.getRange(1, 1, 1, lastCol).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, lastCol).setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      headers.forEach(function(h, i) {
        var col = i + 1;
        var range = sheet.getRange(2, col, lastRow - 1, 1);
        if (/율/.test(h)) range.setNumberFormat('0.0%').setHorizontalAlignment('right');
        else if (/금액|매출|정산|매입|이익|수수료|부가세|공급/.test(h) && !/상품수|건수|수량/.test(h)) range.setNumberFormat('#,##0').setHorizontalAlignment('right');
        else if (/수량|건수|상품수|고객수|주문수|미정산/.test(h)) range.setNumberFormat('#,##0').setHorizontalAlignment('right');
        else if (/날짜|일자/.test(h)) range.setNumberFormat('@').setHorizontalAlignment('center');
        else range.setNumberFormat('@').setHorizontalAlignment('left');
      });
    }
    sheet.setFrozenRows(1);
    for (var c = 1; c <= Math.min(lastCol, 24); c++) sheet.autoResizeColumn(c);
  } catch (e) {}
}

function replaceSheetValuesSafe_v637_(sheet, headers, rows) {
  if (typeof replaceSheetValuesSafe_v628_ === 'function') return replaceSheetValuesSafe_v628_(sheet, headers, rows);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows && rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function splitVatSafe_v637_(amount) {
  if (typeof splitVat_v628_ === 'function') return splitVat_v628_(amount);
  amount = Math.round(num_v637_(amount));
  var supply = Math.round(amount / 1.1);
  return { total: amount, supply: supply, vat: amount - supply };
}

function normalizeAccount_v637_(v) { return String(v == null ? '' : v).replace(/,/g, '').trim(); }
function isValidAccount_v637_(v) { return /^beliun\d{4}$/i.test(String(v || '').trim()); }
function mapAccountNo_v637_(v) { v = normalizeAccount_v637_(v); if (v === '1') return 'beliun1021'; if (v === '2') return 'beliun1024'; if (v === '3') return 'beliun1023'; return ''; }
function num_v637_(v) { if (typeof toNumber_ === 'function') return toNumber_(v); if (typeof v === 'number') return v; var n = Number(String(v == null ? '' : v).replace(/₩|,|%/g, '').trim()); return isNaN(n) ? 0 : n; }
