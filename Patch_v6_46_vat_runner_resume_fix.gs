var LOTTEON_PATCH_V646_VAT_RUNNER_RESUME_FIX_LOADED = true;
var LOTTEON_V646_JOB_KEY = 'LOTTEON_V644_VAT_JOB_STATE';
var LOTTEON_V646_HANDLER = 'generateVatReportsFullSeparated_v622';
var LOTTEON_V646_DELAY_MS = 60000;

generateVatReportsFullSeparated_v622 = function() {
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  var state = getVatStateResume_v646_();

  if (state && state.status === 'running') {
    try { SpreadsheetApp.getActive().toast('부가세 경량 생성 이어실행: ' + (state.phase || '') + ' / offset ' + (state.detailOffset || 0), 'LOTTEON', 5); } catch (e) {}
    return runVatBatchOneTick_v644_(state, true);
  }

  return startVatLightJob_v646_(ui);
};

function startVatLightJob_v646_(ui) {
  clearVatTriggersResume_v646_();
  var state = {
    status: 'running',
    phase: 'vat_detail',
    detailOffset: 0,
    phaseIndex: 0,
    startedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    completed: [],
    errors: [],
    version: 'v6.46',
    chunkSize: 50,
    memo: 'customer/order VAT sheets excluded, A:AC light source, market id column D'
  };
  saveVatStateResume_v646_(state);
  if (typeof writeVatStatus_v644_ === 'function') writeVatStatus_v644_(state, '예약됨', 0, 'v6.46 경량 생성 시작. 다음 실행부터 실제 배치 진행');
  scheduleVatTriggerResume_v646_();
  if (ui) {
    try {
      ui.alert('부가세 신고자료 경량 생성 예약 완료\n\nv6.46 기준으로 기존 상태를 새로 시작했습니다.\n약 1분 뒤 자동 실행됩니다.\n같은 메뉴를 한 번 더 누르면 즉시 다음 배치를 실행합니다.');
    } catch (e) {}
  }
  return { ok: true, scheduled: true, version: 'v6.46', state: state };
}

function getVatStateResume_v646_() {
  var raw = PropertiesService.getScriptProperties().getProperty(LOTTEON_V646_JOB_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function saveVatStateResume_v646_(state) {
  PropertiesService.getScriptProperties().setProperty(LOTTEON_V646_JOB_KEY, JSON.stringify(state));
}

function scheduleVatTriggerResume_v646_() {
  ScriptApp.newTrigger(LOTTEON_V646_HANDLER).timeBased().after(LOTTEON_V646_DELAY_MS).create();
}

function clearVatTriggersResume_v646_() {
  try {
    ScriptApp.getProjectTriggers().forEach(function(t) {
      try {
        if (t.getHandlerFunction && t.getHandlerFunction() === LOTTEON_V646_HANDLER) ScriptApp.deleteTrigger(t);
      } catch (e) {}
    });
  } catch (e) {}
}
