/**
 * LOTTEON v6.24 column width auto-continue patch
 *
 * 목적:
 * - v6.23 열너비 자동조정은 1회 실행당 2개 시트씩 처리하고 사용자가 같은 메뉴를 반복 실행하는 구조였습니다.
 * - v6.24부터는 첫 메뉴 실행 후 남은 시트가 있으면 시간 기반 트리거로 자동 이어실행합니다.
 * - 시간초과 방지를 위해 1회 처리량은 기존처럼 2개 시트로 유지합니다.
 *
 * 운영:
 * - LOTTEON 서식 → 열너비 자동조정 시작/이어실행 1번 실행
 * - 남은 시트가 있으면 약 1분 뒤 자동으로 다음 2개 시트 처리
 * - 완료 시 자동 트리거 정리
 * - 중간에 멈추고 싶으면 열너비 자동조정 초기화 실행
 */

var LOTTEON_PATCH_V624_COLUMN_WIDTH_AUTO_CONTINUE_LOADED = true;
var LOTTEON_COLWIDTH_V624_HANDLER = 'runColumnWidthAutoAdjustStep_v623';
var LOTTEON_COLWIDTH_V624_AUTO_RUNNING = 'LOTTEON_COLWIDTH_V624_AUTO_RUNNING';
var LOTTEON_COLWIDTH_V624_LAST_UPDATED = 'LOTTEON_COLWIDTH_V624_LAST_UPDATED';

runColumnWidthAutoAdjustStep_v623 = function() {
  return runColumnWidthAutoAdjustStepAutoContinue_v624_();
};

resetColumnWidthAutoAdjust_v623 = function() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(LOTTEON_COLWIDTH_V623_NEXT_INDEX);
  props.deleteProperty(LOTTEON_COLWIDTH_V623_LAST_RESULT);
  props.deleteProperty(LOTTEON_COLWIDTH_V624_AUTO_RUNNING);
  props.deleteProperty(LOTTEON_COLWIDTH_V624_LAST_UPDATED);
  deleteColumnWidthAutoContinueTriggers_v624_();
  safeUiAlert_v624_('열너비 자동조정 진행 상태를 초기화했습니다.\n\n예약된 자동 이어실행 트리거도 정리했습니다.');
};

showColumnWidthAutoAdjustStatus_v623 = function() {
  var ss = SpreadsheetApp.getActive();
  var props = PropertiesService.getScriptProperties();
  var sheets = getColumnWidthTargetSheets_v623_(ss);
  var idx = toNumber_(props.getProperty(LOTTEON_COLWIDTH_V623_NEXT_INDEX));
  var autoRunning = props.getProperty(LOTTEON_COLWIDTH_V624_AUTO_RUNNING) === 'Y';
  var last = props.getProperty(LOTTEON_COLWIDTH_V623_LAST_RESULT) || '';
  var updated = props.getProperty(LOTTEON_COLWIDTH_V624_LAST_UPDATED) || '';
  safeUiAlert_v624_(
    '열너비 자동조정 상태\n\n' +
    '대상 시트 수: ' + sheets.length + '\n' +
    '다음 처리 위치: ' + Math.min(idx + 1, sheets.length + 1) + ' / ' + sheets.length + '\n' +
    '자동 이어실행: ' + (autoRunning ? '진행 중' : '대기 없음') + '\n' +
    '최근 갱신: ' + (updated || '없음') + '\n' +
    '최근 결과: ' + (last || '없음')
  );
};

function runColumnWidthAutoAdjustStepAutoContinue_v624_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    scheduleNextColumnWidthStep_v624_();
    safeUiAlert_v624_('다른 작업이 실행 중입니다.\n\n잠시 후 자동으로 다시 이어실행합니다.');
    return { locked: true };
  }

  var started = Date.now();
  try {
    deleteColumnWidthAutoContinueTriggers_v624_();

    var ss = SpreadsheetApp.getActive();
    var props = PropertiesService.getScriptProperties();
    var sheets = getColumnWidthTargetSheets_v623_(ss);
    var startIndex = toNumber_(props.getProperty(LOTTEON_COLWIDTH_V623_NEXT_INDEX));
    if (startIndex < 0 || startIndex >= sheets.length) startIndex = 0;

    var endIndex = Math.min(startIndex + LOTTEON_COLWIDTH_V623_BATCH_SIZE, sheets.length);
    var processed = [];

    for (var i = startIndex; i < endIndex; i++) {
      var sheet = sheets[i];
      normalizeOutputSheetBeforeWidth_v623_(sheet);
      adjustSheetColumnWidthsByData_v623_(sheet);
      processed.push(sheet.getName());
    }

    var elapsedSec = Math.round((Date.now() - started) / 1000);
    var done = endIndex >= sheets.length;
    var result = '처리=' + processed.join(', ') + ' / 위치=' + endIndex + '/' + sheets.length + ' / 소요초=' + elapsedSec;

    props.setProperty(LOTTEON_COLWIDTH_V623_LAST_RESULT, result);
    props.setProperty(LOTTEON_COLWIDTH_V624_LAST_UPDATED, Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'));

    if (done) {
      props.deleteProperty(LOTTEON_COLWIDTH_V623_NEXT_INDEX);
      props.deleteProperty(LOTTEON_COLWIDTH_V624_AUTO_RUNNING);
      deleteColumnWidthAutoContinueTriggers_v624_();
    } else {
      props.setProperty(LOTTEON_COLWIDTH_V623_NEXT_INDEX, String(endIndex));
      props.setProperty(LOTTEON_COLWIDTH_V624_AUTO_RUNNING, 'Y');
      scheduleNextColumnWidthStep_v624_();
    }

    try { log_('column_width_auto_continue_v624', result + ' / done=' + (done ? 'Y' : 'N')); } catch (e) {}

    safeUiAlert_v624_(
      (done ? '열너비 자동조정 완료' : '열너비 자동조정 일부 완료') + '\n\n' +
      '처리 시트:\n- ' + processed.join('\n- ') + '\n\n' +
      '진행: ' + endIndex + ' / ' + sheets.length + '\n' +
      '소요초: ' + elapsedSec + '\n\n' +
      (done ? '전체 대상 시트 처리가 끝났습니다.' : '남은 시트는 약 1분 뒤 자동으로 이어실행됩니다. 같은 메뉴를 계속 누르지 않아도 됩니다.')
    );

    return { done: done, processed: processed, nextIndex: done ? 0 : endIndex };
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    try { log_('column_width_auto_continue_v624_error', msg); } catch (ignore) {}
    safeUiAlert_v624_('열너비 자동조정 중 오류가 발생했습니다.\n\n' + msg + '\n\n상태 확인 후 다시 실행해주세요.');
    throw e;
  } finally {
    lock.releaseLock();
  }
}

function scheduleNextColumnWidthStep_v624_() {
  deleteColumnWidthAutoContinueTriggers_v624_();
  try {
    ScriptApp.newTrigger(LOTTEON_COLWIDTH_V624_HANDLER)
      .timeBased()
      .after(60 * 1000)
      .create();
    PropertiesService.getScriptProperties().setProperty(LOTTEON_COLWIDTH_V624_AUTO_RUNNING, 'Y');
  } catch (e) {
    PropertiesService.getScriptProperties().setProperty(LOTTEON_COLWIDTH_V624_AUTO_RUNNING, 'N');
    try { log_('column_width_auto_continue_trigger_error_v624', String(e)); } catch (ignore) {}
    safeUiAlert_v624_('자동 이어실행 트리거 생성에 실패했습니다.\n\n권한 승인 후 같은 메뉴를 다시 실행해주세요.\n\n' + String(e && e.message ? e.message : e));
  }
}

function deleteColumnWidthAutoContinueTriggers_v624_() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function(t) {
      if (t.getHandlerFunction && t.getHandlerFunction() === LOTTEON_COLWIDTH_V624_HANDLER) {
        ScriptApp.deleteTrigger(t);
      }
    });
  } catch (e) {
    try { log_('column_width_delete_trigger_error_v624', String(e)); } catch (ignore) {}
  }
}

function safeUiAlert_v624_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    // time-driven trigger context에서는 UI alert를 띄울 수 없으므로 무시합니다.
  }
}
