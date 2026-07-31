/**
 * PR #25 production-path operating smoke R12.
 *
 * Uses the current main production bundle (v6.65) and appends only the PR #25
 * v6.66 + v6.67 patches. The normal VAT state-save / 500-row / time-trigger
 * continuation is preserved. No standalone VAT/card reconciliation is used.
 */
const PR25_R12_VERSION = 'v1.23-PR25-PRODUCTION-CONTINUATION-R12';
const PR25_R12_HANDLER = 'continuePr25ProductionVatR12';
const PR25_R12_STATUS_SHEET = 'PR25_실행상태';
const PR25_R12_VAT_STATE_KEY = 'LOTTEON_V648_LIGHT_VAT_JOB_STATE';
const PR25_R12_BRANCH_RAW_BASE = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/codex/issue-24-tracking-payment-primary/';
const PR25_R12_PATCH_URLS = [
  PR25_R12_BRANCH_RAW_BASE + 'Patch_v6_66_vat_tracking_payment_primary.gs',
  PR25_R12_BRANCH_RAW_BASE + 'Patch_v6_67_vat_tracking_production_path_fix.gs'
];

function startPr25ProductionVatR12() {
  pr25r12_assertLoader_();
  pr25r12_clearTriggers_();
  PropertiesService.getScriptProperties().deleteProperty(PR25_R12_VAT_STATE_KEY);
  pr25r12_writeStatus_(SpreadsheetApp.getActive(), null, 'production VAT 초기화 및 첫 배치 시작', 'running');
  return pr25r12_runProduction_('start');
}

function continuePr25ProductionVatR12() {
  pr25r12_assertLoader_();
  return pr25r12_runProduction_('continue');
}

function pr25r12_runProduction_(mode) {
  var ss = pr25r12_resolveSpreadsheet_();
  try {
    var bundle = loadLotteonRemoteBundle_();
    var prPatches = pr25r12_fetchPatchBundle_();
    var invocation = [
      "LOTTEON_PATCH_BOOTSTRAP_VERSION = 'v6.67-PR25';",
      "LOTTEON_V648_HANDLER = '" + PR25_R12_HANDLER + "';",
      'generateVatReportsFullSeparated_v622();'
    ].join('\n');
    var result = eval(bundle + '\n\n;\n\n' + prPatches + '\n\n;\n\n' + invocation);
    var state = pr25r12_getVatState_();
    if (state && state.spreadsheetId) ss = SpreadsheetApp.openById(state.spreadsheetId);
    pr25r12_writeStatus_(ss, state, mode === 'start' ? '첫 배치 예약 완료' : '자동 이어실행 진행', state && state.status || 'running');
    return result;
  } catch (e) {
    var failed = pr25r12_getVatState_() || {};
    failed.status = 'failed';
    failed.phase = failed.phase || 'loader_or_runtime';
    failed.lastError = String(e && e.message ? e.message : e);
    failed.updatedAt = new Date().toISOString();
    pr25r12_writeStatus_(ss, failed, 'production VAT 실행 오류', 'failed');
    pr25r12_clearTriggers_();
    throw e;
  }
}

function pr25r12_resolveSpreadsheet_() {
  var state = pr25r12_getVatState_();
  if (state && state.spreadsheetId) return SpreadsheetApp.openById(state.spreadsheetId);
  return SpreadsheetApp.getActive();
}

function pr25r12_fetchPatchBundle_() {
  var requests = PR25_R12_PATCH_URLS.map(function(url) {
    return { url:url + '?ts=' + new Date().getTime(), method:'get', muteHttpExceptions:true, followRedirects:true };
  });
  var responses = UrlFetchApp.fetchAll(requests);
  var texts = [];
  for (var i = 0; i < responses.length; i++) {
    var code = responses[i].getResponseCode();
    var text = responses[i].getContentText('UTF-8');
    if (code < 200 || code >= 300) {
      throw new Error('PR25 patch load failed HTTP ' + code + ': ' + PR25_R12_PATCH_URLS[i] + '\n' + text.slice(0, 500));
    }
    texts.push(text);
  }
  return texts.join('\n\n;\n\n');
}

function pr25r12_assertLoader_() {
  if (typeof loadLotteonRemoteBundle_ !== 'function') {
    throw new Error('운영 loader의 loadLotteonRemoteBundle_ 함수를 찾지 못했습니다. v1.14-MAIN loader를 유지하세요.');
  }
}

function pr25r12_getVatState_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR25_R12_VAT_STATE_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function pr25r12_clearTriggers_() {
  var handlers = {
    generateVatReportsFullSeparated_v622:true,
    continuePr25ProductionVatR12:true,
    continuePr25TrackingContinuationR6:true,
    continuePr25TrackingContinuationR7:true,
    continuePr25TrackingContinuationR8:true,
    continuePr25TrackingContinuationR9:true,
    continuePr25TrackingContinuationR11:true
  };
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    try {
      if (handlers[trigger.getHandlerFunction()]) ScriptApp.deleteTrigger(trigger);
    } catch (e) {}
  });
}

function pr25r12_writeStatus_(ss, state, message, forcedStatus) {
  state = state || {};
  var sheet = ss.getSheetByName(PR25_R12_STATUS_SHEET) || ss.insertSheet(PR25_R12_STATUS_SHEET);
  var rows = [
    ['항목','값'],
    ['버전',PR25_R12_VERSION],
    ['실행 경로','main production bundle + PR25 v6.66/v6.67'],
    ['상태',forcedStatus || state.status || ''],
    ['단계',state.phase || ''],
    ['메시지',message || ''],
    ['다음 실행 예약',state.nextRunScheduled ? 'Y' : 'N'],
    ['시작시각',state.startedAt || ''],
    ['갱신시각',state.updatedAt || new Date().toISOString()],
    ['원본 진행행',Number(state.sourceRow || 0)],
    ['원본 마지막행',Number(state.sourceLastRow || 0)],
    ['원본 읽기 열수',Number(state.sourceReadColumns || 0)],
    ['트래킹 열번호',Number(state.trackingPaymentColumn || 0)],
    ['fallback 열번호',Number(state.fallbackPaymentColumn || 0)],
    ['부가세 상세 작성행',Number(state.writtenRows || 0)],
    ['제외행',Number(state.skippedRows || 0)],
    ['계정미확인행',Number(state.accountMissingRows || 0)],
    ['마지막 오류',state.lastError || '']
  ];
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange(1, 1, 1, 2).setBackground('#d9eaf7').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 680);
  SpreadsheetApp.flush();
}
