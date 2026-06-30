/**
 * LOTTEON v6.43 VAT launcher status-first auto continue
 *
 * 문제:
 * - v6.42도 첫 실행에서 1단계 작업을 바로 수행한 뒤 상태 시트를 쓰는 구조였습니다.
 * - 1단계가 시간초과되면 부가세_생성상태 시트도 생성되지 않아 진행 여부를 확인할 수 없었습니다.
 *
 * 수정:
 * - 메뉴 첫 실행은 무거운 작업을 하지 않습니다.
 * - 즉시 부가세_생성상태 시트를 만들고, 1단계 실행 트리거만 예약한 뒤 종료합니다.
 * - 예약 트리거 또는 사용자의 재실행 때만 다음 1단계를 처리합니다.
 * - 각 단계 시작 전에 상태 시트를 먼저 갱신합니다.
 * - 따라서 특정 단계가 시간초과되어도 어느 단계에서 멈췄는지 시트에 남습니다.
 */

var LOTTEON_PATCH_V643_VAT_LAUNCHER_STATUS_FIRST_LOADED = true;
var LOTTEON_V643_VAT_JOB_KEY = 'LOTTEON_V642_VAT_JOB_STATE';
var LOTTEON_V643_HANDLER = 'generateVatReportsFullSeparated_v622';
var LOTTEON_V643_DELAY_MS = 60 * 1000;

generateVatReportsFullSeparated_v622 = function() {
  return runVatLauncherStatusFirst_v643_(true);
};

function runVatLauncherStatusFirst_v643_(allowUi) {
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  var state = getVatJobState_v643_();

  if (!state || state.status === 'done' || state.status === 'failed') {
    state = {
      status: 'running',
      stepIndex: 0,
      startedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      completed: [],
      errors: [],
      launcherVersion: 'v6.43',
      mode: 'status_first_auto_continue'
    };
    saveVatJobState_v643_(state);
    writeVatStatusSheet_v643_(state, '예약됨', 0, '첫 실행: 상태 시트 생성 후 1단계 자동 예약');
    clearVatTriggers_v643_();
    scheduleVatTrigger_v643_();
    if (allowUi && ui) {
      safeAlert_v643_(ui,
        '부가세 신고자료 생성 예약 완료\n\n' +
        'v6.43 기준: 첫 실행은 상태 시트만 만들고 무거운 작업은 하지 않습니다.\n' +
        '1단계는 약 1분 뒤 자동 실행됩니다.\n\n' +
        '진행 상태는 부가세_생성상태 시트에서 확인하세요.'
      );
    }
    return { ok: true, scheduled: true, state: state };
  }

  return runVatOneStep_v643_(state, allowUi);
}

function runVatOneStep_v643_(state, allowUi) {
  var started = Date.now();
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  var steps = getVatSteps_v643_();
  var step = steps[state.stepIndex];

  if (!step) {
    state.status = 'done';
    state.updatedAt = new Date().toISOString();
    saveVatJobState_v643_(state);
    clearVatTriggers_v643_();
    writeVatStatusSheet_v643_(state, '완료', Date.now() - started, '모든 단계 완료');
    if (allowUi && ui) safeAlert_v643_(ui, buildVatDoneMessage_v643_(state));
    return { ok: true, done: true, state: state };
  }

  try {
    state.currentStep = { index: state.stepIndex, key: step.key, label: step.label, startedAt: new Date().toISOString() };
    state.updatedAt = new Date().toISOString();
    saveVatJobState_v643_(state);
    writeVatStatusSheet_v643_(state, '실행중', 0, '현재 단계 실행 중: ' + step.label);

    var salesAgg = prepareVatSalesAgg_v643_();
    step.fn(salesAgg);

    state.completed = state.completed || [];
    state.completed.push({ index: state.stepIndex, key: step.key, label: step.label, completedAt: new Date().toISOString() });
    state.stepIndex++;
    state.currentStep = null;
    state.updatedAt = new Date().toISOString();

    var done = state.stepIndex >= steps.length;
    if (done) {
      state.status = 'done';
      saveVatJobState_v643_(state);
      clearVatTriggers_v643_();
      writeVatStatusSheet_v643_(state, '완료', Date.now() - started, '모든 단계 완료');
      if (allowUi && ui) safeAlert_v643_(ui, buildVatDoneMessage_v643_(state));
      return { ok: true, done: true, state: state };
    }

    saveVatJobState_v643_(state);
    clearVatTriggers_v643_();
    scheduleVatTrigger_v643_();
    writeVatStatusSheet_v643_(state, '진행중', Date.now() - started, '다음 단계 자동 예약됨');

    if (allowUi && ui) {
      safeAlert_v643_(ui,
        '부가세 신고자료 생성 진행중\n\n' +
        '완료 단계: ' + state.stepIndex + ' / ' + steps.length + '\n' +
        '방금 완료: ' + step.label + '\n' +
        '다음 단계: ' + steps[state.stepIndex].label + '\n\n' +
        '다음 단계는 약 1분 뒤 자동 이어실행됩니다.'
      );
    }
    return { ok: true, done: false, state: state };
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    state.status = 'failed';
    state.updatedAt = new Date().toISOString();
    state.errors = state.errors || [];
    state.errors.push({ stepIndex: state.stepIndex, stepLabel: step.label, message: msg, at: new Date().toISOString() });
    saveVatJobState_v643_(state);
    clearVatTriggers_v643_();
    writeVatStatusSheet_v643_(state, '오류', Date.now() - started, msg);
    try { log_('patch_v643_vat_step_error', 'step=' + step.label + ' message=' + msg); } catch (ignore) {}
    if (allowUi && ui) safeAlert_v643_(ui, '부가세 신고자료 생성 중 오류가 발생했습니다.\n\n단계: ' + step.label + '\n오류: ' + msg + '\n\n부가세_생성상태 시트를 확인하세요.');
    throw e;
  }
}

function getVatSteps_v643_() {
  if (typeof LOTTEON_V642_VAT_STEPS !== 'undefined' && LOTTEON_V642_VAT_STEPS && LOTTEON_V642_VAT_STEPS.length) return LOTTEON_V642_VAT_STEPS;
  return [
    { key: 'vat_detail', label: '부가세_신고자료', fn: function(agg) { buildVatDetailSingleSource_v628_(agg); } },
    { key: 'vat_product', label: '부가세_상품별', fn: function(agg) { buildVatProductSingleSource_v628_(agg); } },
    { key: 'vat_customer', label: '부가세_고객별', fn: function(agg) { buildVatCustomerSingleSource_v628_(agg); } },
    { key: 'vat_order', label: '부가세_주문번호별', fn: function(agg) { buildVatOrderSingleSource_v628_(agg); } },
    { key: 'account_summary', label: '사업자별_계정별_써머리', fn: function(agg) { if (typeof buildAccountSummarySheet_v637_ === 'function') buildAccountSummarySheet_v637_(agg); } },
    { key: 'final_format', label: '최종 최소 서식', fn: function(agg) { if (typeof applyVatAutoContinueFinalFormat_v642_ === 'function') applyVatAutoContinueFinalFormat_v642_(); } }
  ];
}

function prepareVatSalesAgg_v643_() {
  if (typeof prepareVatAutoContinueSalesAgg_v642_ === 'function') return prepareVatAutoContinueSalesAgg_v642_();
  var salesAgg = buildSingleSourceSalesAgg_v628_();
  if (typeof forceSalesAggAccountFromColumnD_v638_ === 'function') salesAgg = forceSalesAggAccountFromColumnD_v638_(salesAgg);
  if (typeof applyCustomerAddressGroups_v629_ === 'function') salesAgg = applyCustomerAddressGroups_v629_(salesAgg);
  if (typeof recalculateProfitRates_v632_ === 'function') salesAgg = recalculateProfitRates_v632_(salesAgg);
  return salesAgg;
}

function getVatJobState_v643_() {
  var raw = PropertiesService.getScriptProperties().getProperty(LOTTEON_V643_VAT_JOB_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function saveVatJobState_v643_(state) {
  PropertiesService.getScriptProperties().setProperty(LOTTEON_V643_VAT_JOB_KEY, JSON.stringify(state));
}

function scheduleVatTrigger_v643_() {
  ScriptApp.newTrigger(LOTTEON_V643_HANDLER).timeBased().after(LOTTEON_V643_DELAY_MS).create();
}

function clearVatTriggers_v643_() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    try {
      if (t.getHandlerFunction && t.getHandlerFunction() === LOTTEON_V643_HANDLER) ScriptApp.deleteTrigger(t);
    } catch (e) {}
  });
}

function writeVatStatusSheet_v643_(state, statusText, elapsedMs, memo) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('부가세_생성상태') || ss.insertSheet('부가세_생성상태');
  var steps = getVatSteps_v643_();
  var next = state.stepIndex < steps.length ? steps[state.stepIndex].label : '없음';
  var current = state.currentStep && state.currentStep.label ? state.currentStep.label : '';
  var rows = [
    ['항목','값','메모'],
    ['상태', statusText, 'v6.43 상태 먼저 생성'],
    ['진행단계', (state.stepIndex || 0) + ' / ' + steps.length, '완료된 단계 수'],
    ['현재단계', current, '실행 중인 단계'],
    ['다음단계', next, ''],
    ['시작시각', state.startedAt || '', ''],
    ['갱신시각', new Date().toISOString(), ''],
    ['이번 실행 소요초', Math.round((elapsedMs || 0) / 1000), ''],
    ['완료단계', (state.completed || []).map(function(x){ return x.label; }).join(' → '), ''],
    ['오류', (state.errors || []).map(function(x){ return '[' + x.stepLabel + '] ' + x.message; }).join(' / '), ''],
    ['메모', memo || '', '']
  ];
  try {
    sheet.clearContents();
    sheet.getRange(1, 1, rows.length, 3).setValues(rows);
    sheet.getRange(1, 1, 1, 3).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.autoResizeColumns(1, 3);
  } catch (e) {}
}

function buildVatDoneMessage_v643_(state) {
  var steps = getVatSteps_v643_();
  return '부가세 신고자료 생성 완료\n\n' +
    '기준: v6.43 상태 먼저 생성 + 자동 이어실행\n' +
    '완료 단계: ' + steps.length + ' / ' + steps.length + '\n\n' +
    '확인 시트: 부가세_신고자료, 사업자별_계정별_써머리, 사업자번호_매핑검증, 부가세_생성상태';
}

function safeAlert_v643_(ui, message) {
  try { ui.alert(message); } catch (e) {}
}
