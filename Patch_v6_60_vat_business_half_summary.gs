/** v6.60 Issue #20: prepend business/payment half-year VAT filing summary to 부가세_기간별. */
var LOTTEON_PATCH_V660_VAT_BUSINESS_HALF_SUMMARY_LOADED = true;

/*
 * Keep the raw payment/card-company label from 매출데이터_붙여넣기.
 * No issuer inference: values such as 신한카드, 우리카드(일시불), L.PAY신용카드,
 * 카카오페이 머니 are preserved exactly as supplied.
 */
var __baseVatHeaderIndexes_v660_ = typeof vatHeaderIndexes_v648_ === 'function' ? vatHeaderIndexes_v648_ : null;
var __baseVatDetailHeaders_v660_ = typeof vatDetailHeaders_v648_ === 'function' ? vatDetailHeaders_v648_ : null;
var __baseVatDetailRow_v660_ = typeof vatDetailRow_v648_ === 'function' ? vatDetailRow_v648_ : null;

if (__baseVatHeaderIndexes_v660_) {
  vatHeaderIndexes_v648_ = function(headers) {
    var indexes = __baseVatHeaderIndexes_v660_.apply(this, arguments) || {};
    indexes.payment = findVatPaymentHeader_v660_(headers || []);
    return indexes;
  };
}

if (__baseVatDetailHeaders_v660_) {
  vatDetailHeaders_v648_ = function() {
    var headers = __baseVatDetailHeaders_v660_.apply(this, arguments).slice();
    if (headers.indexOf('결제수단') < 0) headers.push('결제수단');
    return headers;
  };
}

if (__baseVatDetailRow_v660_) {
  vatDetailRow_v648_ = function(row, ix, sourceRow) {
    var result = __baseVatDetailRow_v660_.apply(this, arguments);
    if (!result || !result.row) return result;
    var payment = '';
    if (ix && typeof ix.payment === 'number' && ix.payment >= 0) {
      payment = cleanVatText_v648_(valueAt_v648_(row, ix.payment));
    }
    result.row.push(payment || '결제수단 미확인');
    return result;
  };
}

function findVatPaymentHeader_v660_(headers) {
  var names = ['결제수단','결제정보','결제방법','카드사','결제수단/카드사','결제수단(카드사)'];
  for (var n = 0; n < names.length; n++) {
    var wanted = String(names[n]).replace(/\s/g, '');
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i] == null ? '' : headers[i]).replace(/\s/g, '') === wanted) return i;
    }
  }
  return -1;
}

var __baseBuildVatPeriodSummary_v660_ = typeof buildVatPeriodSummary_v657_ === 'function' ? buildVatPeriodSummary_v657_ : null;

if (__baseBuildVatPeriodSummary_v660_) {
  buildVatPeriodSummary_v657_ = function(ss) {
    ss = ss || SpreadsheetApp.getActive();
    var result = __baseBuildVatPeriodSummary_v660_.apply(this, arguments);
    var rendered = prependVatBusinessHalfSummary_v660_(ss);
    if (result && typeof result === 'object') result.businessHalfSummaryRows = rendered.summaryRows;
    return result;
  };
}

function prependVatBusinessHalfSummary_v660_(ss) {
  var detail = ss && ss.getSheetByName && ss.getSheetByName('부가세_신고자료');
  var periodSheet = ss && ss.getSheetByName && ss.getSheetByName('부가세_기간별');
  if (!detail || !periodSheet || detail.getLastRow() < 1 || periodSheet.getLastRow() < 1) {
    return { summaryRows: 0, reason: 'MISSING_SHEET' };
  }

  var detailValues = detail.getDataRange().getValues();
  var periodValues = periodSheet.getDataRange().getValues();
  var summary = aggregateVatBusinessHalf_v660_(detailValues);
  var summaryHeaders = vatBusinessHalfHeaders_v660_();
  var detailStartRow = summary.length + 5;

  periodSheet.clearContents();
  periodSheet.getRange(1, 1).setValue('사업자별 반기 신고요약 (결제수단/카드사별)');
  periodSheet.getRange(2, 1, 1, summaryHeaders.length).setValues([summaryHeaders]);
  if (summary.length) periodSheet.getRange(3, 1, summary.length, summaryHeaders.length).setValues(summary);
  if (periodValues.length && periodValues[0].length) {
    periodSheet.getRange(detailStartRow, 1, periodValues.length, periodValues[0].length).setValues(periodValues);
  }

  formatVatBusinessHalfSummary_v660_(periodSheet, summaryHeaders, summary.length, detailStartRow, periodValues[0] || []);
  periodSheet.setFrozenRows(2);
  return { summaryRows: summary.length, detailStartRow: detailStartRow };
}

function vatBusinessHalfHeaders_v660_() {
  return [
    '신고연도','반기','사업자등록번호','결제수단/카드사','연결 쿠팡계정ID','주문건수',
    '순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료',
    '매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','비고'
  ];
}

function aggregateVatBusinessHalf_v660_(values) {
  if (!values || values.length < 2) return [];
  var headers = values[0] || [];
  var ix = function(names) { return findVatPeriodHeader_v657_(headers, names, -1); };
  var yearIx = ix(['신고연도']);
  var halfIx = ix(['반기']);
  var accountIx = ix(['쿠팡계정ID']);
  var businessIx = ix(['사업자등록번호']);
  var paymentIx = ix(['결제수단','결제수단/카드사','카드사','결제정보','결제방법']);
  var orderIx = ix(['주문번호','마켓주문번호','주문ID','주문ID(마켓)']);
  var salesIx = ix(['순수매출액']);
  var salesSupplyIx = ix(['매출공급가액']);
  var salesVatIx = ix(['매출부가세']);
  var settlementIx = ix(['정산기준금액']);
  var feeIx = ix(['마켓수수료/비용','마켓수수료']);
  var purchaseIx = ix(['매입금액']);
  var purchaseSupplyIx = ix(['매입공급가액']);
  var purchaseVatIx = ix(['매입부가세']);
  var payableIx = ix(['납부예상부가세']);
  var profitIx = ix(['예상이익']);
  var vatProfitIx = ix(['부가세반영예상이익']);
  var required = [yearIx, halfIx, accountIx, businessIx, paymentIx, salesIx, salesSupplyIx, salesVatIx, settlementIx, feeIx, purchaseIx, purchaseSupplyIx, purchaseVatIx, payableIx, profitIx, vatProfitIx];
  if (required.some(function(index) { return index < 0; })) throw new Error('v6.60 사업자/결제수단별 반기 신고요약 필수 헤더를 찾을 수 없습니다.');

  var map = {};
  function num(row, index) { return index < 0 ? 0 : vatPeriodNumber_v657_(row[index]); }
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var year = String(row[yearIx] == null ? '' : row[yearIx]).trim();
    var half = String(row[halfIx] == null ? '' : row[halfIx]).trim();
    if (!year || year === '기간미확인' || (half !== '상반기' && half !== '하반기')) continue;

    var account = String(row[accountIx] == null ? '' : row[accountIx]).trim();
    var rawBusiness = String(row[businessIx] == null ? '' : row[businessIx]).trim();
    var business = rawBusiness || '사업자번호 미매핑';
    var rawPayment = String(row[paymentIx] == null ? '' : row[paymentIx]).trim();
    var payment = rawPayment || '결제수단 미확인';
    var key = [year, half, business, payment].join('|');
    if (!map[key]) {
      map[key] = {
        year: year, half: half, business: business, payment: payment, accounts: {}, orders: {}, blankOrderRows: 0,
        sales: 0, salesSupply: 0, salesVat: 0, settlement: 0, fee: 0,
        purchase: 0, purchaseSupply: 0, purchaseVat: 0, payable: 0, profit: 0, vatProfit: 0
      };
    }
    var item = map[key];
    if (account) item.accounts[account] = true;
    var orderNo = orderIx >= 0 ? String(row[orderIx] == null ? '' : row[orderIx]).trim() : '';
    if (orderNo) item.orders[orderNo] = true; else item.blankOrderRows++;
    item.sales += num(row, salesIx);
    item.salesSupply += num(row, salesSupplyIx);
    item.salesVat += num(row, salesVatIx);
    item.settlement += num(row, settlementIx);
    item.fee += num(row, feeIx);
    item.purchase += num(row, purchaseIx);
    item.purchaseSupply += num(row, purchaseSupplyIx);
    item.purchaseVat += num(row, purchaseVatIx);
    item.payable += num(row, payableIx);
    item.profit += num(row, profitIx);
    item.vatProfit += num(row, vatProfitIx);
  }

  var halfRank = { '상반기': 0, '하반기': 1 };
  return Object.keys(map).map(function(key) {
    var item = map[key];
    var notes = [];
    if (item.business === '사업자번호 미매핑') notes.push('사업자번호 미매핑');
    if (item.payment === '결제수단 미확인') notes.push('결제수단 미확인');
    if (item.blankOrderRows) notes.push('주문번호 공란 ' + item.blankOrderRows + '행');
    return [
      item.year, item.half, item.business, item.payment, Object.keys(item.accounts).sort().join(', '), Object.keys(item.orders).length,
      item.sales, item.salesSupply, item.salesVat, item.settlement, item.fee,
      item.purchase, item.purchaseSupply, item.purchaseVat, item.payable, item.profit, item.vatProfit, notes.join(' / ')
    ];
  }).sort(function(a, b) {
    return String(a[0]).localeCompare(String(b[0])) ||
      (halfRank[a[1]] - halfRank[b[1]]) ||
      String(a[2]).localeCompare(String(b[2])) ||
      String(a[3]).localeCompare(String(b[3]));
  });
}

function formatVatBusinessHalfSummary_v660_(sheet, summaryHeaders, summaryRows, detailStartRow, detailHeaders) {
  var maxColumns = Math.max(summaryHeaders.length, detailHeaders.length || 0, 1);
  sheet.getRange(1, 1, 1, maxColumns).setBackground('#b4c6e7').setFontWeight('bold');
  sheet.getRange(2, 1, 1, summaryHeaders.length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
  if (detailHeaders && detailHeaders.length) {
    sheet.getRange(detailStartRow, 1, 1, detailHeaders.length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
  }

  var summaryPlan = vatBusinessHalfFormatPlan_v660_(summaryHeaders);
  Object.keys(summaryPlan).forEach(function(key) {
    if (summaryRows > 0) sheet.getRange(3, Number(key) + 1, summaryRows, 1).setNumberFormat(summaryPlan[key]);
  });

  var detailPlan = typeof vatPeriodFormatPlan_v657_ === 'function' ? vatPeriodFormatPlan_v657_(detailHeaders || []) : {};
  Object.keys(detailPlan).forEach(function(key) {
    var detailRows = Math.max(sheet.getLastRow() - detailStartRow, 0);
    if (detailRows > 0) sheet.getRange(detailStartRow + 1, Number(key) + 1, detailRows, 1).setNumberFormat(detailPlan[key]);
  });

  for (var i = 0; i < maxColumns; i++) {
    var header = summaryHeaders[i] || detailHeaders[i] || '';
    var normalized = String(header).replace(/\s/g, '');
    var width = /순수매출액|공급가액|부가세|정산기준금액|마켓수수료|매입금액|예상이익/.test(normalized) ? 115 :
      (/계정ID|사업자등록번호|결제수단|카드사/.test(normalized) ? 145 : 95);
    sheet.setColumnWidth(i + 1, width);
  }
}

function vatBusinessHalfFormatPlan_v660_(headers) {
  var plan = {};
  headers.forEach(function(header, index) {
    var h = String(header || '').replace(/\s/g, '');
    if (/주문건수|순수매출액|매출공급가액|매출부가세|정산기준금액|마켓수수료|매입금액|매입공급가액|매입부가세|납부예상부가세|예상이익|부가세반영예상이익/.test(h)) plan[index] = '#,##0';
    else if (/신고연도|반기|사업자등록번호|계정ID|결제수단|카드사/.test(h)) plan[index] = '@';
  });
  return plan;
}
