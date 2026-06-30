/**
 * LOTTEON v6.42 VAT auto-continue step execution
 *
 * 문제:
 * - v6.41에서 중복 생성은 줄였지만, 한 번의 실행 안에서 모든 부가세/계정/검증 시트를 생성하면
 *   Apps Script 최대 실행 시간 초과가 계속 발생할 수 있습니다.
 *
 * 수정:
 * - 자동 열너비 메뉴처럼 부가세 신고자료 생성 작업을 단계별로 쪼개서 실행합니다.
 * - 메뉴를 한 번 누르면 상태를 PropertiesService에 저장하고, ScriptApp 시간 트리거로 다음 단계를 자동 이어실행합니다.
 * - 각 실행은 한 단계만 처리합니다.
 * - trigger handler는 기존 로더 함수 generateVatReportsFullSeparated_v622를 그대로 사용합니다.
 * - 이전 단계가 끝난 뒤 다음 단계가 예약되므로, 전체 작업이 한 번에 시간초과되지 않습니다.
 */

var LOTTEON_PATCH_V642_VAT_AUTO_CONTINUE_STEPS_LOADED = true;
var LOTTEON_V642_VAT_JOB_KEY = 'LOTTEON_V642_VAT_JOB_STATE';
var LOTTEON_V642_VAT_HANDLER = 'generateVatReportsFullSeparated_v622';
var LOTTEON_V642_VAT_TRIGGER_DELAY_MS = 60 * 1000;

var LOTTEON_V642_VAT_STEPS = [
  { key: 'vat_detail', label: '부가세_신고자료', fn: function(agg) { buildVatDetailSingleSource_v628_(agg); } },
  { key: 'vat_product', label: '부가세_상품별', fn: function(agg) { buildVatProductSingleSource_v628_(agg); } },
  { key: 'vat_customer', label: '부가세_고객별', fn: function(agg) { buildVatCustomerSingleSource_v628_(agg); } },
  { key: 'vat_order', label: '부가세_주문번호별', fn: function(agg) { buildVatOrderSingleSource_v628_(agg); } },
  { key: 'brand_margin', label: '브랜드별_마진율', fn: function(agg) { if (typeof buildBrandMarginSingleSource_v628_ === 'function') buildBrandMarginSingleSource_v628_(agg); } },
  { key: 'account_summary', label: '사업자별_계정별_써머리', fn: function(agg) { if (typeof buildAccountSummarySheet_v637_ === 'function') buildAccountSummarySheet_v637_(agg); } },
  { key: 'unsettled', label: '미정산_쿠팡계정별', fn: function(agg) { if (typeof buildUnsettledSettlementByAccountSheet_v629_ === 'function') buildUnsettledSettlementByAccountSheet_v629_(agg); } },
  { key: 'validation_amounts', label: '시트별_금액검증', fn: function(agg) { if (typeof buildFinancialValidationSheets_v628_ === 'function') buildFinancialValidationSheets_v628_(agg); } },
  { key: 'diag_account', label: '계정/사업자번호 검증', fn: function(agg) {
      if (typeof buildMarketIdAccountDiagnostic_v638_ === 'function') buildMarketIdAccountDiagnostic_v638_(agg);
      if (typeof buildBusinessNoMappingDiagnostic_v640_ === 'function') buildBusinessNoMappingDiagnostic_v640_(agg);
    }
  },
  { key: 'diag_customer', label: '고객주소_구분검증', fn: function(agg) { if (typeof buildCustomerAddressGroupingDiagnostics_v629_ === 'function') buildCustomerAddressGroupingDiagnostics_v629_(agg); } },
  { key: 'final_format', label: '최종 최소 서식', fn: function(agg) { applyVatAutoContinueFinalFormat_v642_(); } }
];

generateVatReportsFullSeparated_v622 = function() {
  return runVatAutoContinueStep_v642_(true);
};

function runVatAutoContinueStep_v642_(allowUi) {
  var started = Date.now();
  var state = getVatJobState_v642_();
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}

  if (!state || state.status === 'done' || state.status === 'failed') {
    state = {
      status: 'running',
      stepIndex: 0,
      startedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      completed: [],
      errors: []
    };
    saveVatJobState_v642_(state);
    clearVatAutoContinueTriggers_v642_();
  }

  try {
    var step = LOTTEON_V642_VAT_STEPS[state.stepIndex];
    if (!step) {
      state.status = 'done';
      state.updatedAt = new Date().toISOString();
      saveVatJobState_v642_(state);
      clearVatAutoContinueTriggers_v642_();
      writeVatAutoContinueStatusSheet_v642_(state, '완료', Date.now() - started);
      if (allowUi && ui) safeAlert_v642_(ui, buildVatDoneMessage_v642_(state));
      return { ok: true, done: true, state: state };
    }

    var salesAgg = prepareVatAutoContinueSalesAgg_v642_();
    step.fn(salesAgg);

    state.completed.push({ index: state.stepIndex, key: step.key, label: step.label, completedAt: new Date().toISOString() });
    state.stepIndex++;
    state.updatedAt = new Date().toISOString();

    var done = state.stepIndex >= LOTTEON_V642_VAT_STEPS.length;
    if (done) {
      state.status = 'done';
      saveVatJobState_v642_(state);
      clearVatAutoContinueTriggers_v642_();
      writeVatAutoContinueStatusSheet_v642_(state, '완료', Date.now() - started);
      if (allowUi && ui) safeAlert_v642_(ui, buildVatDoneMessage_v642_(state));
      return { ok: true, done: true, state: state };
    }

    saveVatJobState_v642_(state);
    scheduleVatAutoContinueTrigger_v642_();
    writeVatAutoContinueStatusSheet_v642_(state, '진행중', Date.now() - started);

    if (allowUi && ui) {
      safeAlert_v642_(ui,
        '부가세 신고자료 생성 진행중\n\n' +
        '완료 단계: ' + state.stepIndex + ' / ' + LOTTEON_V642_VAT_STEPS.length + '\n' +
        '방금 완료: ' + step.label + '\n' +
        '다음 단계: ' + LOTTEON_V642_VAT_STEPS[state.stepIndex].label + '\n\n' +
        '다음 단계는 약 1분 뒤 자동 이어실행됩니다.\n' +
        '상태는 부가세_생성상태 시트에서 확인하세요.'
      );
    }
    return { ok: true, done: false, state: state };
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    state.status = 'failed';
    state.updatedAt = new Date().toISOString();
    state.errors = state.errors || [];
    state.errors.push({ stepIndex: state.stepIndex, message: msg, at: new Date().toISOString() });
    saveVatJobState_v642_(state);
    clearVatAutoContinueTriggers_v642_();
    writeVatAutoContinueStatusSheet_v642_(state, '오류', Date.now() - started);
    try { log_('patch_v642_vat_auto_continue_error', msg); } catch (ignore) {}
    if (allowUi && ui) safeAlert_v642_(ui, '부가세 신고자료 생성 중 오류가 발생했습니다.\n\n' + msg + '\n\n부가세_생성상태 시트를 확인하세요.');
    throw e;
  }
}

function prepareVatAutoContinueSalesAgg_v642_() {
  var salesAgg = buildSingleSourceSalesAgg_v628_();
  if (typeof forceSalesAggAccountFromColumnD_v638_ === 'function') salesAgg = forceSalesAggAccountFromColumnD_v638_(salesAgg);
  if (typeof applyCustomerAddressGroups_v629_ === 'function') salesAgg = applyCustomerAddressGroups_v629_(salesAgg);
  if (typeof recalculateProfitRates_v632_ === 'function') salesAgg = recalculateProfitRates_v632_(salesAgg);
  return salesAgg;
}

function applyVatAutoContinueFinalFormat_v642_() {
  var ss = SpreadsheetApp.getActive();
  if (typeof formatVatBusinessNoSheet_v640_ === 'function') formatVatBusinessNoSheet_v640_(ss.getSheetByName('부가세_신고자료'));
  if (typeof formatAccountOutputSheet_v637_ === 'function') formatAccountOutputSheet_v637_(ss.getSheetByName('사업자별_계정별_써머리'));
  if (typeof hardFixBrandMarginRateColumns_v634_ === 'function') hardFixBrandMarginRateColumns_v634_();
  if (typeof sortUnsettledSheetByOrderDateAsc_v635_ === 'function') sortUnsettledSheetByOrderDateAsc_v635_();
}

function getVatJobState_v642_() {
  var raw = PropertiesService.getScriptProperties().getProperty(LOTTEON_V642_VAT_JOB_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function saveVatJobState_v642_(state) {
  PropertiesService.getScriptProperties().setProperty(LOTTEON_V642_VAT_JOB_KEY, JSON.stringify(state));
}

function scheduleVatAutoContinueTrigger_v642_() {
  clearVatAutoContinueTriggers_v642_();
  ScriptApp.newTrigger(LOTTEON_V642_VAT_HANDLER).timeBased().after(LOTTEON_V642_VAT_TRIGGER_DELAY_MS).create();
}

function clearVatAutoContinueTriggers_v642_() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    try {
      if (t.getHandlerFunction && t.getHandlerFunction() === LOTTEON_V642_VAT_HANDLER) ScriptApp.deleteTrigger(t);
    } catch (e) {}
  });
}

function writeVatAutoContinueStatusSheet_v642_(state, statusText, elapsedMs) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('부가세_생성상태') || ss.insertSheet('부가세_생성상태');
  var rows = [
    ['항목','값','메모'],
    ['상태', statusText, 'v6.42 자동 이어실행'],
    ['진행단계', (state.stepIndex || 0) + ' / ' + LOTTEON_V642_VAT_STEPS.length, '완료된 단계 수'],
    ['다음단계', state.stepIndex < LOTTEON_V642_VAT_STEPS.length ? LOTTEON_V642_VAT_STEPS[state.stepIndex].label : '없음', ''],
    ['시작시각', state.startedAt || '', ''],
    ['갱신시각', state.updatedAt || '', ''],
    ['이번 실행 소요초', Math.round((elapsedMs || 0) / 1000), ''],
    ['완료단계', (state.completed || []).map(function(x){ return x.label; }).join(' → '), ''],
    ['오류', (state.errors || []).map(function(x){ return x.message; }).join(' / '), '']
  ];
  try {
    sheet.clearContents();
    sheet.getRange(1,1,rows.length,3).setValues(rows);
    sheet.getRange(1,1,1,3).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.autoResizeColumns(1,3);
  } catch (e) {}
}

function buildVatDoneMessage_v642_(state) {
  return '부가세 신고자료 생성 완료\n\n' +
    '기준: v6.42 자동 이어실행\n' +
    '완료 단계: ' + LOTTEON_V642_VAT_STEPS.length + ' / ' + LOTTEON_V642_VAT_STEPS.length + '\n\n' +
    '확인 시트:\n' +
    '- 부가세_신고자료\n' +
    '- 사업자별_계정별_써머리\n' +
    '- 사업자번호_매핑검증\n' +
    '- 부가세_생성상태';
}

function safeAlert_v642_(ui, message) {
  try { ui.alert(message); } catch (e) {}
}
