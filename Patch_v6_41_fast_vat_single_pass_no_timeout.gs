/**
 * LOTTEON v6.41 fast VAT single-pass no-timeout patch
 *
 * 문제:
 * - 부가세 신고자료 생성(무거움) 메뉴가 시간초과됩니다.
 * - v6.37~v6.40에서 generateVatReportsFullSeparated_v622를 여러 patch가 계속 감싸면서
 *   같은 원본 집계/시트 생성을 반복 실행하는 구조가 되었습니다.
 * - 데이터 생성 작업인데 base generate chain을 모두 호출하면 부가세 시트, 계정별 시트, 검증 시트가 중복 재생성되어 시간이 늘어납니다.
 *
 * 수정:
 * - generateVatReportsFullSeparated_v622를 마지막 단계에서 완전히 재정의합니다.
 * - 이전 base generate를 호출하지 않습니다.
 * - buildSingleSourceSalesAgg_v628_는 1회만 실행합니다.
 * - D열 마켓아이디 계정 구분, 사업자번호 매핑을 적용한 뒤 필요한 출력 시트만 1회씩 생성합니다.
 */

var LOTTEON_PATCH_V641_FAST_VAT_SINGLE_PASS_NO_TIMEOUT_LOADED = true;

generateVatReportsFullSeparated_v622 = function() {
  var ui = SpreadsheetApp.getUi();
  var started = Date.now();
  try {
    var salesAgg = buildSingleSourceSalesAgg_v628_();

    if (typeof forceSalesAggAccountFromColumnD_v638_ === 'function') {
      salesAgg = forceSalesAggAccountFromColumnD_v638_(salesAgg);
    }
    if (typeof applyCustomerAddressGroups_v629_ === 'function') {
      salesAgg = applyCustomerAddressGroups_v629_(salesAgg);
    }
    if (typeof recalculateProfitRates_v632_ === 'function') {
      salesAgg = recalculateProfitRates_v632_(salesAgg);
    }

    // 핵심 출력 시트만 1회씩 생성합니다. 이전 generate wrapper chain은 호출하지 않습니다.
    buildVatDetailSingleSource_v628_(salesAgg);
    buildVatProductSingleSource_v628_(salesAgg);
    buildVatCustomerSingleSource_v628_(salesAgg);
    buildVatOrderSingleSource_v628_(salesAgg);

    if (typeof buildBrandMarginSingleSource_v628_ === 'function') buildBrandMarginSingleSource_v628_(salesAgg);
    if (typeof buildAccountSummarySheet_v637_ === 'function') buildAccountSummarySheet_v637_(salesAgg);
    if (typeof buildUnsettledSettlementByAccountSheet_v629_ === 'function') buildUnsettledSettlementByAccountSheet_v629_(salesAgg);

    // 검증/진단 시트도 1회만 생성합니다.
    if (typeof buildFinancialValidationSheets_v628_ === 'function') buildFinancialValidationSheets_v628_(salesAgg);
    if (typeof buildMarketIdAccountDiagnostic_v638_ === 'function') buildMarketIdAccountDiagnostic_v638_(salesAgg);
    if (typeof buildBusinessNoMappingDiagnostic_v640_ === 'function') buildBusinessNoMappingDiagnostic_v640_(salesAgg);
    if (typeof buildCustomerAddressGroupingDiagnostics_v629_ === 'function') buildCustomerAddressGroupingDiagnostics_v629_(salesAgg);

    // 무거운 전체 autoResize/전체시트 서식 정리는 호출하지 않습니다.
    // 필요한 시트만 최소 서식 적용합니다.
    if (typeof formatVatBusinessNoSheet_v640_ === 'function') formatVatBusinessNoSheet_v640_(SpreadsheetApp.getActive().getSheetByName('부가세_신고자료'));
    if (typeof hardFixBrandMarginRateColumns_v634_ === 'function') hardFixBrandMarginRateColumns_v634_();
    if (typeof applyPercentOneDecimalToSheet_v633_ === 'function') {
      ['브랜드별_마진율','사업자별_계정별_써머리','대시보드'].forEach(function(name) {
        var sh = SpreadsheetApp.getActive().getSheetByName(name);
        if (sh) applyPercentOneDecimalToSheet_v633_(sh);
      });
    }

    var elapsed = Math.round((Date.now() - started) / 1000);
    try { log_('patch_v641_fast_vat_single_pass', 'rows=' + (salesAgg.detailRows || []).length + ' elapsed=' + elapsed); } catch (e) {}

    ui.alert(
      '부가세 신고자료 생성 완료\n\n' +
      '기준: v6.41 단일 실행 / 중복 생성 방지\n' +
      '분석대상 행수: ' + ((salesAgg.detailRows || []).length) + '\n' +
      '순수매출액: ' + fmtNum_v641_(salesAgg.totalNetSales) + '\n' +
      '매입금액: ' + fmtNum_v641_(salesAgg.totalPurchase) + '\n' +
      '소요초: ' + elapsed + '\n\n' +
      '확인 시트: 부가세_신고자료, 사업자별_계정별_써머리, 사업자번호_매핑검증'
    );
    return { ok: true, rows: (salesAgg.detailRows || []).length, elapsedSeconds: elapsed };
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    try { log_('patch_v641_fast_vat_single_pass_error', msg); } catch (ignore) {}
    ui.alert('부가세 신고자료 생성 중 오류가 발생했습니다.\n\n' + msg);
    throw e;
  }
};

function fmtNum_v641_(n) {
  n = Math.round(Number(n || 0));
  return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}
