/**
 * LOTTEON v6.45 light VAT generation without customer/order sheets
 *
 * 확정 기준:
 * - 부가세_고객별, 부가세_주문번호별 시트는 더 이상 만들지 않습니다.
 * - 기존 시트가 있으면 삭제합니다.
 * - 부가세 생성은 매출데이터_붙여넣기 전체 폭을 읽지 않고 A:AC(29열)까지만 읽습니다.
 * - 가장 무거운 부가세_신고자료는 50행씩 배치 생성합니다.
 * - 고객주소 구분 작업도 부가세 생성 경로에서는 제외합니다.
 */

var LOTTEON_PATCH_V645_LIGHT_VAT_NO_CUSTOMER_ORDER_LOADED = true;
var LOTTEON_V645_DETAIL_CHUNK_SIZE = 50;
var LOTTEON_V645_MAX_SOURCE_COL = 29; // AC

// v6.44 배치 크기를 더 작게 덮어씁니다.
if (typeof LOTTEON_V644_DETAIL_CHUNK_SIZE !== 'undefined') LOTTEON_V644_DETAIL_CHUNK_SIZE = LOTTEON_V645_DETAIL_CHUNK_SIZE;

// 메뉴 실행 시 기존 꼬인 상태를 v6.45 작업으로 새로 시작합니다.
generateVatReportsFullSeparated_v622 = function() {
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  deleteDeprecatedVatSheets_v645_();

  var state = {
    status: 'running',
    phase: 'vat_detail',
    detailOffset: 0,
    phaseIndex: 0,
    startedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    completed: [],
    errors: [],
    version: 'v6.45',
    chunkSize: LOTTEON_V645_DETAIL_CHUNK_SIZE,
    memo: '부가세_고객별/부가세_주문번호별 제외, A:AC 경량 집계'
  };

  PropertiesService.getScriptProperties().setProperty('LOTTEON_V644_VAT_JOB_STATE', JSON.stringify(state));
  if (typeof saveVatState_v644_ === 'function') saveVatState_v644_(state);
  if (typeof clearVatTriggers_v644_ === 'function') clearVatTriggers_v644_();
  if (typeof writeVatStatus_v644_ === 'function') writeVatStatus_v644_(state, '예약됨', 0, 'v6.45 경량 생성 시작: 고객별/주문번호별 제외, 50행 배치');
  if (typeof scheduleVatTrigger_v644_ === 'function') scheduleVatTrigger_v644_();

  if (ui) {
    try {
      ui.alert(
        '부가세 신고자료 경량 생성 예약 완료\n\n' +
        'v6.45 기준:\n' +
        '- 부가세_고객별 제외\n' +
        '- 부가세_주문번호별 제외\n' +
        '- 매출데이터_붙여넣기 A:AC만 읽음\n' +
        '- 부가세_신고자료 50행씩 배치 생성\n\n' +
        '진행 상태는 부가세_생성상태 시트에서 확인하세요.'
      );
    } catch (e) {}
  }
  return { ok: true, scheduled: true, version: 'v6.45' };
};

// v6.44가 매 배치마다 호출하는 집계 함수를 경량 버전으로 교체합니다.
prepareVatSalesAgg_v644_ = function() {
  return buildSingleSourceSalesAggLight_v645_();
};

// v6.44 나머지 단계에서 고객별/주문번호별/고객주소 검증을 제거합니다.
getOtherVatSteps_v644_ = function() {
  return [
    { key: 'vat_product', label: '부가세_상품별', fn: function(agg) { buildVatProductSingleSource_v628_(agg); } },
    { key: 'brand_margin', label: '브랜드별_마진율', fn: function(agg) { if (typeof buildBrandMarginSingleSource_v628_ === 'function') buildBrandMarginSingleSource_v628_(agg); } },
    { key: 'account_summary', label: '사업자별_계정별_써머리', fn: function(agg) { if (typeof buildAccountSummarySheet_v637_ === 'function') buildAccountSummarySheet_v637_(agg); } },
    { key: 'unsettled', label: '미정산_쿠팡계정별', fn: function(agg) { if (typeof buildUnsettledSettlementByAccountSheet_v629_ === 'function') buildUnsettledSettlementByAccountSheet_v629_(agg); } },
    { key: 'validation_amounts', label: '시트별_금액검증', fn: function(agg) { if (typeof buildFinancialValidationSheets_v628_ === 'function') buildFinancialValidationSheets_v628_(agg); } },
    { key: 'diag_account', label: '계정/사업자번호 검증', fn: function(agg) { if (typeof buildMarketIdAccountDiagnostic_v638_ === 'function') buildMarketIdAccountDiagnostic_v638_(agg); if (typeof buildBusinessNoMappingDiagnostic_v640_ === 'function') buildBusinessNoMappingDiagnostic_v640_(agg); } },
    { key: 'final_format', label: '최종 최소 서식', fn: function(agg) { applyVatFinalFormatLight_v645_(); } }
  ];
};

function buildSingleSourceSalesAggLight_v645_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SALES_IN) || ss.getSheetByName('매출데이터_붙여넣기');
  if (!sheet || sheet.getLastRow() < 2) throw new Error('매출데이터_붙여넣기 시트를 찾지 못했거나 데이터가 없습니다.');

  var lastRow = sheet.getLastRow();
  var maxCol = Math.min(sheet.getLastColumn(), LOTTEON_V645_MAX_SOURCE_COL);
  var values = sheet.getRange(1, 1, lastRow, maxCol).getValues();
  var headers = values[0].map(function(h) { return String(h || '').trim(); });
  var acIdx = LOTTEON_V628_PURCHASE_AC_COL - 1;
  if (acIdx >= maxCol) throw new Error('AC열 구매가격을 읽을 수 없습니다. 현재 읽은 열 수: ' + maxCol);

  var agg = initSalesAgg_v628_ ? initSalesAgg_v628_(sheet.getName(), headers[acIdx] || 'AC열') : { detailRows: [], byBrand: {}, productSet: {}, orderSet: {}, unsettledOrderSet: {}, overdueUnsettledOrderSet: {} };
  var targetRows = [];
  var excludedRows = [];

  for (var r = 1; r < values.length; r++) {
    var rowObj = rowToObject_v628_(headers, values[r]);
    var sales = num_v628_(getAny_v628_(rowObj, ['결제금액합계(원)','결제금액합계','결제금액','매출액','판매금액']));
    var purchase = num_v628_(values[r][acIdx]);
    var row = makeSourceRow_v628_(r + 1, rowObj, sales, purchase);

    // D열 마켓아이디를 계정 단일 원천으로 강제합니다.
    row.accountId = String(values[r][3] == null ? '' : values[r][3]).replace(/,/g, '').trim() || '마켓아이디없음';

    agg.rawRowCount = (agg.rawRowCount || 0) + 1;
    agg.totalGrossSales = (agg.totalGrossSales || 0) + sales;
    agg.totalRawPurchase = (agg.totalRawPurchase || 0) + purchase;

    var exclude = getAnalysisExclusionReason_v628_(row);
    if (exclude) {
      row.excludeReason = exclude;
      excludedRows.push(row);
      if (row.cancelExcluded) {
        agg.cancelRowCount = (agg.cancelRowCount || 0) + 1;
        agg.totalCancelSales = (agg.totalCancelSales || 0) + sales;
        agg.totalCancelPurchase = (agg.totalCancelPurchase || 0) + purchase;
      } else {
        agg.analysisExcludedRowCount = (agg.analysisExcludedRowCount || 0) + 1;
        agg.totalAnalysisExcludedSales = (agg.totalAnalysisExcludedSales || 0) + sales;
        agg.totalAnalysisExcludedPurchase = (agg.totalAnalysisExcludedPurchase || 0) + purchase;
      }
      continue;
    }
    targetRows.push(row);
  }

  agg.excludedRows = excludedRows;
  agg.targetSourceRows = targetRows;
  targetRows.forEach(function(row) {
    var d = sourceRowToDetail_v628_(row);
    d.accountId = row.accountId || d.accountId || '마켓아이디없음';
    addDetailRowToAgg_v628_(agg, d);
  });

  finalizeSalesAgg_v628_(agg);
  agg.basisVersion = 'v6.45 light A:AC no customer/order VAT';
  try { log_('single_source_agg_light_v645', 'raw=' + agg.rawRowCount + ' target=' + agg.detailRows.length); } catch (e) {}
  return agg;
}

function deleteDeprecatedVatSheets_v645_() {
  var ss = SpreadsheetApp.getActive();
  ['부가세_고객별', '부가세_주문번호별', '고객주소_구분검증'].forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (sh) {
      try { ss.deleteSheet(sh); } catch (e) { try { sh.hideSheet(); } catch (ignore) {} }
    }
  });
}

function applyVatFinalFormatLight_v645_() {
  var ss = SpreadsheetApp.getActive();
  deleteDeprecatedVatSheets_v645_();
  if (typeof formatVatBusinessNoSheet_v640_ === 'function') formatVatBusinessNoSheet_v640_(ss.getSheetByName('부가세_신고자료'));
  if (typeof formatAccountOutputSheet_v637_ === 'function') formatAccountOutputSheet_v637_(ss.getSheetByName('사업자별_계정별_써머리'));
  if (typeof hardFixBrandMarginRateColumns_v634_ === 'function') hardFixBrandMarginRateColumns_v634_();
  if (typeof sortUnsettledSheetByOrderDateAsc_v635_ === 'function') sortUnsettledSheetByOrderDateAsc_v635_();
}
