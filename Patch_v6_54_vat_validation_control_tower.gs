/**
 * LOTTEON v6.54 - Issue #5 VAT generation/validation control tower.
 * Loaded last so the loader-facing VAT entrypoint cannot resume an older job.
 */
var LOTTEON_PATCH_V654_VAT_VALIDATION_CONTROL_TOWER_LOADED = true;
var LOTTEON_V654_VERSION = 'v6.54';
var LOTTEON_V654_LEGACY_JOB_KEYS = [
  'LOTTEON_V642_VAT_JOB_STATE',
  'LOTTEON_V644_VAT_JOB_STATE',
  'LOTTEON_V647_DIRECT_VAT_JOB_STATE',
  'LOTTEON_V648_LIGHT_VAT_JOB_STATE'
];

LOTTEON_V648_VERSION = LOTTEON_V654_VERSION;

generateVatReportsFullSeparated_v622 = function() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: false, busy: true };
  try {
    var state = getVatState_v648_();
    if (!state || state.version !== LOTTEON_V654_VERSION || state.status === 'done' || state.status === 'failed') {
      resetLegacyVatJobState_v654_();
      return startVatJob_v648_();
    }
    return continueVatJob_v648_(state);
  } finally {
    lock.releaseLock();
  }
};

function resetLegacyVatJobState_v654_() {
  var props = PropertiesService.getScriptProperties();
  LOTTEON_V654_LEGACY_JOB_KEYS.forEach(function(key) { props.deleteProperty(key); });
  clearVatTriggers_v648_();
}

vatHeaderIndexes_v648_ = function(headers) {
  var normalized = {};
  headers.forEach(function(header, index) { normalized[cleanVatText_v648_(header).replace(/\s/g, '')] = index; });
  function find(names, fallback) {
    for (var i = 0; i < names.length; i++) {
      var key = names[i].replace(/\s/g, '');
      if (Object.prototype.hasOwnProperty.call(normalized, key)) return normalized[key];
    }
    return fallback;
  }
  return {
    date: find(['마켓주문일자','주문일자','결제일자','주문일시'], 0),
    orderNo: find(['마켓주문번호','주문번호','주문ID','주문ID(마켓)'], 2),
    marketAccount: find(['마켓아이디','쿠팡계정ID','계정ID'], 3),
    marketProductNo: find(['마켓상품번호'], 4),
    sales: find(['결제금액합계(원)','결제금액합계','결제금액','순수매출액','판매금액'], 6),
    marketStatus: find(['마켓주문상태','주문상태','클레임상태','처리상태'], 15),
    themangoStatus: find(['더망고주문상태','더망고상태'], -1),
    customer: find(['고객명','수령인명','수령인','수취인','구매자','주문자'], 8),
    siteName: find(['구매사이트명','구매사이트','사이트명'], 21),
    siteProductNo: find(['사이트상품번호','상품번호','상품코드','판매자상품코드'], 22),
    brand: find(['브랜드','브랜드명'], 23),
    productName: find(['마켓상품명','상품명(원문)','상품명','상품명(옵션포함)','등록상품명'], 16),
    quantity: find(['결제수량','판매수량','수량','구매수량'], 18),
    settlement: find(['정산예정금액(원)','정산예정금액','실제정산금액','정산금액'], -1)
  };
};

vatDetailRow_v648_ = function(row, ix, sourceRow) {
  var accountId = cleanVatText_v648_(valueAt_v648_(row, ix.marketAccount >= 0 ? ix.marketAccount : 3));
  var accountMissing = !accountId;
  var status = [valueAt_v648_(row, ix.marketStatus), valueAt_v648_(row, ix.themangoStatus)].join(' ');
  if (isExcludedOrderStatus_v654_(status)) return { row: null, accountMissing: accountMissing, excluded: 'status' };

  var siteName = cleanVatText_v648_(valueAt_v648_(row, ix.siteName));
  if (!/lotteon/i.test(siteName)) return { row: null, accountMissing: accountMissing, excluded: 'site' };
  var sales = vatNumber_v648_(valueAt_v648_(row, ix.sales));
  if (!sales) return { row: null, accountMissing: accountMissing, excluded: 'sales' };
  var brand = cleanVatText_v648_(valueAt_v648_(row, ix.brand));
  if (!brand) return { row: null, accountMissing: accountMissing, excluded: 'brand' };
  var productNo = cleanVatText_v648_(valueAt_v648_(row, ix.siteProductNo)) || cleanVatText_v648_(valueAt_v648_(row, ix.marketProductNo));
  if (!productNo) return { row: null, accountMissing: accountMissing, excluded: 'product' };

  var settlementActual = vatNumber_v648_(valueAt_v648_(row, ix.settlement));
  var settlement = settlementActual || Math.round(sales * 0.901);
  var purchase = vatNumber_v648_(row[28]);
  var salesVat = splitVat_v648_(sales);
  var purchaseVat = splitVat_v648_(purchase);
  var fee = sales - settlement;
  var profit = settlement - purchase;
  var payableVat = salesVat.vat - purchaseVat.vat;
  var businessNo = businessNoForMarketId_v648_(accountId);

  return { accountMissing: accountMissing, row: [
    vatDateText_v648_(valueAt_v648_(row, ix.date)), accountId, businessNo,
    cleanVatText_v648_(valueAt_v648_(row, ix.orderNo)), cleanVatText_v648_(valueAt_v648_(row, ix.customer)),
    brand, productNo, cleanVatText_v648_(valueAt_v648_(row, ix.productName)),
    vatNumber_v648_(valueAt_v648_(row, ix.quantity)) || 1,
    sales, salesVat.supply, salesVat.vat, settlement, fee, purchase,
    purchaseVat.supply, purchaseVat.vat, payableVat, profit, profit - payableVat,
    accountMissing ? 'D열 마켓아이디 공란 (원본 ' + sourceRow + '행)' : (businessNo ? LOTTEON_V654_VERSION + ' 분석대상 기준' : '사업자번호 미매핑')
  ] };
};

function isExcludedOrderStatus_v654_(status) {
  return /취소|반품|교환|환불|cancel|return|exchange|refund/i.test(String(status || ''));
}

businessNoForMarketId_v648_ = function(marketId) {
  var normalized = cleanVatText_v648_(marketId).toLowerCase();
  if (normalized === 'beliun1021' || normalized === '1021') return '227-27-04928';
  if (normalized === 'beliun1021-1' || normalized === '1021-1') return '176-71-00758';
  if (normalized === 'beliun1023' || normalized === '1023') return '835-58-00765';
  if (normalized === 'beliun1024' || normalized === '1024') return '606-45-93763';
  return '';
};

buildVatProductSummary_v648_ = function(ss, rows) {
  var map = {};
  rows.forEach(function(r) {
    var key = cleanVatText_v648_(r[6]) || cleanVatText_v648_(r[7]) || '상품미확인';
    if (!map[key]) map[key] = [key, r[7], r[5], 0, 0, 0, 0, 0, 0, 0, 0];
    map[key][3] += vatNumber_v648_(r[8]);
    map[key][4] += vatNumber_v648_(r[9]);
    map[key][5] += vatNumber_v648_(r[10]);
    map[key][6] += vatNumber_v648_(r[11]);
    map[key][7] += vatNumber_v648_(r[12]);
    map[key][8] += vatNumber_v648_(r[14]);
    map[key][9] += vatNumber_v648_(r[17]);
    map[key][10] += vatNumber_v648_(r[19]);
  });
  var out = ss.getSheetByName('부가세_상품별') || ss.insertSheet('부가세_상품별');
  writeTable_v648_(out,
    ['상품번호','상품명','브랜드명','판매수량','순수매출액','매출공급가액','매출부가세','정산기준금액','매입금액','납부예상부가세','부가세반영예상이익'],
    objectRows_v648_(map));
};

finishVatSummaries_v648_ = function(ss, state) {
  var detail = ss.getSheetByName(LOTTEON_V648_DETAIL_SHEET);
  var rows = detail && detail.getLastRow() > 1
    ? detail.getRange(2, 1, detail.getLastRow() - 1, vatDetailHeaders_v648_().length).getValues()
    : [];
  buildVatProductSummary_v648_(ss, rows);
  buildBrandMarginSummary_v648_(ss, rows);
  buildAccountSummary_v648_(ss, rows);
  buildMappingDiagnostic_v648_(ss, rows);
  buildSheetAmountValidation_v654_(ss);

  state.status = 'done';
  state.phase = 'done';
  state.nextRunScheduled = false;
  state.updatedAt = new Date().toISOString();
  state.lastError = '';
  saveVatState_v648_(state);
  clearVatTriggers_v648_();
  writeVatStatus_v648_(ss, state, 'v6.54 생성/검증 완료');
  toastVat_v648_(ss, '부가세 자료 및 시트별 금액검증 생성이 완료되었습니다.');
  return { ok: true, done: true, state: state };
};

function buildSheetAmountValidation_v654_(ss) {
  var detail = readTableAmounts_v628_(ss.getSheetByName('부가세_신고자료'));
  var headers = ['시트명','기준 순수매출액','시트 순수매출액','매출 차이','기준 매입금액','시트 매입금액','매입 차이','기준 정산기준금액','시트 정산기준금액','정산 차이','판정','메모'];
  var targets = [
    ['대시보드','summary'],
    ['브랜드별_마진율','table'],
    ['부가세_신고자료','table'],
    ['부가세_상품별','table']
  ];
  var rows = targets.map(function(target) {
    var actual = target[1] === 'summary'
      ? readDashboardSummaryAmounts_v628_(ss.getSheetByName(target[0]))
      : readTableAmounts_v628_(ss.getSheetByName(target[0]));
    var salesDiff = Math.round(detail.sales - actual.sales);
    var purchaseDiff = Math.round(detail.purchase - actual.purchase);
    var settlementDiff = Math.round(detail.settlement - actual.settlement);
    var ok = salesDiff === 0 && purchaseDiff === 0 && settlementDiff === 0;
    return [target[0], detail.sales, actual.sales, salesDiff, detail.purchase, actual.purchase, purchaseDiff,
      detail.settlement, actual.settlement, settlementDiff, ok ? 'OK' : '차이확인', actual.memo || ''];
  });
  var sheet = ss.getSheetByName('시트별_금액검증') || ss.insertSheet('시트별_금액검증');
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground('#d9eaf7').setFontWeight('bold');
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.setFrozenRows(1);
}
