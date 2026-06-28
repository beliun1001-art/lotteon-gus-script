/**
 * LOTTEON v6.22 fast dashboard skip VAT reports patch
 *
 * 문제:
 * - v6.21 이후 ⑤ 대시보드만 빠른 갱신에서 부가세_신고자료/상품별/고객별/주문번호별까지 같이 재생성되어
 *   행 수가 많은 부가세 상세 시트 처리 중 최대 실행 시간 초과가 발생했습니다.
 *
 * 수정:
 * - ⑤ 대시보드만 빠른 갱신은 대시보드 + 브랜드별_마진율 + 미정산_쿠팡계정별 중심으로만 실행합니다.
 * - 부가세 상세 4개 시트는 ⑤에서 자동 생성하지 않습니다.
 * - 부가세 상세 생성은 generateVatReportsFullSeparated_v622() 함수로 분리합니다.
 *
 * 운영:
 * - 평소: ⑤ 대시보드만 빠른 갱신
 * - 신고 전/필요 시: Apps Script 함수 선택에서 generateVatReportsFullSeparated_v622 실행
 */

var LOTTEON_PATCH_V622_FAST_DASHBOARD_SKIP_VAT_REPORTS_LOADED = true;

refreshDashboardFastOnly = function() {
  var ui = SpreadsheetApp.getUi();
  var started = Date.now();
  try {
    var salesAgg = aggregateSalesByBrand_v611_();
    var filterAgg = aggregateFiltersByBrand_v611_();

    // ① 대시보드 재작성: 부가세 상세 생성 없이 운영 요약만 갱신
    if (typeof rebuildDashboardBrandBased_v611_ === 'function') {
      rebuildDashboardBrandBased_v611_(salesAgg, filterAgg);
    }

    // ② 브랜드별 마진율: v6.21 기준 매출/매입 부가세 및 부가세반영예상이익 포함
    if (typeof buildBrandMarginSheet_v611_ === 'function') {
      buildBrandMarginSheet_v611_(salesAgg);
    }

    // ③ 미정산 계정별: 30일초과 주문번호 상세 중심
    if (typeof buildUnsettledSettlementByAccountSheet_v616_ === 'function') {
      buildUnsettledSettlementByAccountSheet_v616_(salesAgg);
    }

    // ④ 빠른 서식만 적용. 전체 시트 상세 축약/재작성은 하지 않음.
    if (typeof applyFastOutputSheetFormatting_v620_ === 'function') {
      applyFastOutputSheetFormatting_v620_();
    }

    var elapsedSec = Math.round((Date.now() - started) / 1000);
    try {
      log_('dashboard_fast_only_v622', 'elapsedSec=' + elapsedSec + ' / vatReports=skipped');
    } catch (e) {}

    ui.alert(
      '대시보드 빠른 갱신 완료\n\n' +
      '소요초: ' + elapsedSec + '\n' +
      '갱신 시트:\n' +
      '- 대시보드\n' +
      '- 브랜드별_마진율\n' +
      '- 미정산_쿠팡계정별\n\n' +
      '시간초과 방지를 위해 부가세 상세 시트는 이번 ⑤에서 제외했습니다.\n' +
      '부가세 상세가 필요하면 generateVatReportsFullSeparated_v622 함수를 별도로 실행하세요.'
    );
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    try { log_('dashboard_fast_only_v622_error', msg); } catch (ignore) {}
    ui.alert('대시보드 빠른 갱신 중 오류가 발생했습니다.\n\n' + msg);
    throw e;
  }
};

function generateVatReportsFullSeparated_v622() {
  var ui = SpreadsheetApp.getUi();
  var started = Date.now();
  try {
    var salesAgg = aggregateSalesByBrand_v611_();

    if (typeof buildVatBreakdownSheetsVatCredit_v621_ === 'function') {
      buildVatBreakdownSheetsVatCredit_v621_(salesAgg);
    } else if (typeof buildVatReportSheet_v611_ === 'function') {
      buildVatReportSheet_v611_(salesAgg);
    } else {
      throw new Error('부가세 신고자료 생성 함수를 찾지 못했습니다.');
    }

    if (typeof applyFastOutputSheetFormatting_v620_ === 'function') {
      applyFastOutputSheetFormatting_v620_();
    }

    var elapsedSec = Math.round((Date.now() - started) / 1000);
    try { log_('vat_reports_full_v622', 'elapsedSec=' + elapsedSec); } catch (e) {}

    ui.alert(
      '부가세 신고자료 생성 완료\n\n' +
      '소요초: ' + elapsedSec + '\n' +
      '갱신 시트:\n' +
      '- 부가세_신고자료\n' +
      '- 부가세_상품별\n' +
      '- 부가세_고객별\n' +
      '- 부가세_주문번호별'
    );
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    try { log_('vat_reports_full_v622_error', msg); } catch (ignore) {}
    ui.alert('부가세 신고자료 생성 중 오류가 발생했습니다.\n\n' + msg);
    throw e;
  }
}

// 과거 buildVatReportSheet_v611_를 호출하는 내부 흐름이 ⑤와 섞일 경우를 방지하기 위한 경량 안전장치.
// generateVatReportsFullSeparated_v622에서 직접 상세 생성 함수를 호출하므로, ⑤ 경로에서는 이 함수가 무거운 상세를 만들지 않게 합니다.
buildVatReportSheet_v611_ = function(salesAgg) {
  try { log_('vat_report_skipped_v622', 'Use generateVatReportsFullSeparated_v622 for full VAT report.'); } catch (e) {}
  return { skipped: true, reason: 'v6.22: VAT reports separated from fast dashboard refresh' };
};
