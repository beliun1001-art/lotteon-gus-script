/**
 * LOTTEON v6.40 business registration number VAT field
 *
 * 확정 기준:
 * - 사업자번호는 매출데이터_붙여넣기 D열 마켓아이디 기준 계정에 매핑합니다.
 * - 1021 / beliun1021 → 227-27-04928
 * - 1023 / beliun1023 → 835-58-00765
 * - 1024 / beliun1024 → 606-45-93763
 * - 부가세_신고자료 시트에 사업자등록번호 열을 추가합니다.
 * - 위치는 날짜, 쿠팡계정ID 다음입니다.
 */

var LOTTEON_PATCH_V640_BUSINESS_REGISTRATION_NO_VAT_FIELD_LOADED = true;

var LOTTEON_V640_BIZ_NO_BY_ACCOUNT = {
  '1021': '227-27-04928',
  'beliun1021': '227-27-04928',
  '1023': '835-58-00765',
  'beliun1023': '835-58-00765',
  '1024': '606-45-93763',
  'beliun1024': '606-45-93763'
};

var __baseBuildVatDetailSingleSource_v640 = typeof buildVatDetailSingleSource_v628_ === 'function' ? buildVatDetailSingleSource_v628_ : null;
var __baseGenerateVatReportsFullSeparated_v640 = typeof generateVatReportsFullSeparated_v622 === 'function' ? generateVatReportsFullSeparated_v622 : null;
var __baseApplyDisplayStandardsOnlyFast_v640 = typeof applyDisplayStandardsOnlyFast_v623 === 'function' ? applyDisplayStandardsOnlyFast_v623 : null;
var __baseApplyFastOutputSheetFormatting_v640 = typeof applyFastOutputSheetFormatting_v620_ === 'function' ? applyFastOutputSheetFormatting_v620_ : null;

buildVatDetailSingleSource_v628_ = function(salesAgg) {
  salesAgg = salesAgg || (typeof buildSingleSourceSalesAgg_v628_ === 'function' ? buildSingleSourceSalesAgg_v628_() : null);
  if (typeof forceSalesAggAccountFromColumnD_v638_ === 'function') salesAgg = forceSalesAggAccountFromColumnD_v638_(salesAgg);
  if (typeof prepareAccountSummarySalesAgg_v637_ === 'function') salesAgg = prepareAccountSummarySalesAgg_v637_(salesAgg);

  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('부가세_신고자료') || ss.insertSheet('부가세_신고자료');
  var headers = [
    '날짜','쿠팡계정ID','사업자등록번호','주문번호','고객명','브랜드명','상품번호','상품명','판매수량',
    '순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료/비용',
    '매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','비고'
  ];
  var rows = (salesAgg && salesAgg.detailRows || []).map(function(d) { return vatDetailRowWithBusinessNo_v640_(d); });
  replaceSheetValuesSafe_v640_(sheet, headers, rows);
  formatVatBusinessNoSheet_v640_(sheet);
  buildBusinessNoMappingDiagnostic_v640_(salesAgg);
  return { rows: rows.length };
};

generateVatReportsFullSeparated_v622 = function() {
  var result = __baseGenerateVatReportsFullSeparated_v640 ? __baseGenerateVatReportsFullSeparated_v640.apply(this, arguments) : null;
  try {
    var salesAgg = typeof buildSingleSourceSalesAgg_v628_ === 'function' ? buildSingleSourceSalesAgg_v628_() : null;
    if (typeof forceSalesAggAccountFromColumnD_v638_ === 'function') salesAgg = forceSalesAggAccountFromColumnD_v638_(salesAgg);
    buildVatDetailSingleSource_v628_(salesAgg);
    buildBusinessNoMappingDiagnostic_v640_(salesAgg);
  } catch (e) { try { log_('patch_v640_after_vat_error', String(e && e.message ? e.message : e)); } catch(ignore) {} }
  return result || { ok: true };
};

applyDisplayStandardsOnlyFast_v623 = function() {
  var result = __baseApplyDisplayStandardsOnlyFast_v640 ? __baseApplyDisplayStandardsOnlyFast_v640.apply(this, arguments) : null;
  try {
    var salesAgg = typeof buildSingleSourceSalesAgg_v628_ === 'function' ? buildSingleSourceSalesAgg_v628_() : null;
    if (typeof forceSalesAggAccountFromColumnD_v638_ === 'function') salesAgg = forceSalesAggAccountFromColumnD_v638_(salesAgg);
    ensureVatDetailBusinessNoField_v640_(salesAgg);
    SpreadsheetApp.getUi().alert('표시서식 정리 완료\n\n부가세_신고자료 사업자등록번호 열을 확인했습니다.');
  } catch (e) {}
  return result || { ok: true };
};

applyFastOutputSheetFormatting_v620_ = function() {
  var result = __baseApplyFastOutputSheetFormatting_v640 ? __baseApplyFastOutputSheetFormatting_v640.apply(this, arguments) : null;
  try { formatVatBusinessNoSheet_v640_(SpreadsheetApp.getActive().getSheetByName('부가세_신고자료')); } catch (e) {}
  return result || { ok: true };
};

function vatDetailRowWithBusinessNo_v640_(d) {
  var salesSplit = splitVatSafe_v640_(d.netSales);
  var purchaseSplit = splitVatSafe_v640_(d.purchase);
  var payableVat = salesSplit.vat - purchaseSplit.vat;
  var accountId = String(d.accountId || '계정미확인').trim();
  var bizNo = getBusinessNoByAccount_v640_(accountId);
  return [
    d.dateText || '',
    accountId,
    bizNo,
    d.orderNo || '',
    d.customer || '',
    d.brand || '',
    d.productNo || '',
    d.productName || '',
    num_v640_(d.qty),
    num_v640_(d.netSales),
    salesSplit.supply,
    salesSplit.vat,
    num_v640_(d.settlementBasis),
    num_v640_(d.marketFee),
    num_v640_(d.purchase),
    purchaseSplit.supply,
    purchaseSplit.vat,
    payableVat,
    num_v640_(d.estimatedProfit),
    num_v640_(d.estimatedProfit) - payableVat,
    bizNo ? 'v6.40 사업자번호 포함' : '사업자번호 매핑 확인 필요'
  ];
}

function ensureVatDetailBusinessNoField_v640_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('부가세_신고자료');
  if (!sheet || sheet.getLastRow() < 1) return buildVatDetailSingleSource_v628_(salesAgg);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h || '').replace(/\s+/g, '').trim(); });
  if (headers.indexOf('사업자등록번호') >= 0) {
    fillExistingVatBusinessNoColumn_v640_(sheet);
    return { ok: true, already: true };
  }
  return buildVatDetailSingleSource_v628_(salesAgg || (typeof buildSingleSourceSalesAgg_v628_ === 'function' ? buildSingleSourceSalesAgg_v628_() : null));
}

function fillExistingVatBusinessNoColumn_v640_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || '').replace(/\s+/g, '').trim(); });
  var accountIdx = headers.indexOf('쿠팡계정ID');
  var bizIdx = headers.indexOf('사업자등록번호');
  if (accountIdx < 0 || bizIdx < 0) return;
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (var r = 0; r < values.length; r++) {
    values[r][bizIdx] = getBusinessNoByAccount_v640_(values[r][accountIdx]);
  }
  sheet.getRange(2, 1, values.length, lastCol).setValues(values);
  formatVatBusinessNoSheet_v640_(sheet);
}

function buildBusinessNoMappingDiagnostic_v640_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('사업자번호_매핑검증') || ss.insertSheet('사업자번호_매핑검증');
  var counts = {};
  var missing = 0;
  (salesAgg && salesAgg.detailRows || []).forEach(function(d) {
    var account = String(d.accountId || '계정미확인').trim();
    var biz = getBusinessNoByAccount_v640_(account);
    var key = account + '|' + (biz || '사업자번호없음');
    counts[key] = (counts[key] || 0) + 1;
    if (!biz) missing++;
  });
  var rows = Object.keys(counts).sort().map(function(k) {
    var parts = k.split('|');
    return [parts[0], parts[1], counts[k]];
  });
  rows.unshift(['beliun1024 / 1024', '606-45-93763', '기준 매핑']);
  rows.unshift(['beliun1023 / 1023', '835-58-00765', '기준 매핑']);
  rows.unshift(['beliun1021 / 1021', '227-27-04928', '기준 매핑']);
  rows.unshift(['사업자번호 미매핑 행수', missing, '0이면 정상']);
  try {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 3).setValues([['계정/항목','사업자등록번호','행수/메모']]);
    if (rows.length) sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    sheet.getRange(1, 1, 1, 3).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.autoResizeColumns(1, 3);
  } catch (e) {}
}

function getBusinessNoByAccount_v640_(accountId) {
  var raw = String(accountId == null ? '' : accountId).replace(/,/g, '').trim();
  if (!raw) return '';
  if (LOTTEON_V640_BIZ_NO_BY_ACCOUNT[raw]) return LOTTEON_V640_BIZ_NO_BY_ACCOUNT[raw];
  var m = raw.match(/1021|1023|1024/);
  if (m && LOTTEON_V640_BIZ_NO_BY_ACCOUNT[m[0]]) return LOTTEON_V640_BIZ_NO_BY_ACCOUNT[m[0]];
  return '';
}

function formatVatBusinessNoSheet_v640_(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || '').replace(/\n/g, '').replace(/\s+/g, '').trim(); });
  try {
    sheet.getRange(1, 1, 1, lastCol).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    if (lastRow > 1) {
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
    sheet.autoResizeColumns(1, Math.min(lastCol, 21));
  } catch (e) {}
}

function replaceSheetValuesSafe_v640_(sheet, headers, rows) {
  if (typeof replaceSheetValuesSafe_v628_ === 'function') return replaceSheetValuesSafe_v628_(sheet, headers, rows);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows && rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function splitVatSafe_v640_(amount) {
  if (typeof splitVatSafe_v637_ === 'function') return splitVatSafe_v637_(amount);
  if (typeof splitVat_v628_ === 'function') return splitVat_v628_(amount);
  amount = Math.round(num_v640_(amount));
  var supply = Math.round(amount / 1.1);
  return { total: amount, supply: supply, vat: amount - supply };
}

function num_v640_(v) {
  if (typeof toNumber_ === 'function') return toNumber_(v);
  if (typeof v === 'number') return v;
  var n = Number(String(v == null ? '' : v).replace(/₩|,|%/g, '').trim());
  return isNaN(n) ? 0 : n;
}
