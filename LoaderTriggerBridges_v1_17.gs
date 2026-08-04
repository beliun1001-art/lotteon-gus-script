/**
 * LOTTEON Apps Script permanent trigger bridges v1.17
 *
 * One-time local install. Future temporary batch runners are updated only in
 * GitHub's permanent `remote/task-runner/RemoteTaskCurrent.gs` slot.
 *
 * v1.17:
 * - Preserve v1.16 permanent remote start/continue bridge.
 * - Declare MailApp scope locally so terminal PASS/ERROR/ROLLBACK notices can
 *   be authorized and delivered without manual sheet inspection.
 */
const LOTTEON_TRIGGER_BRIDGE_VERSION = 'v1.17';
const LOTTEON_REMOTE_TASK_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/remote/task-runner/RemoteTaskCurrent.gs';
const LOTTEON_REMOTE_NOTICE_EMAIL = 'beliun1001@gmail.com';

/** Existing local entrypoint used by the daily 06:10 filter scheduler. */
function runDailyFilterCountsStart() {
  return runFilterCountLightweightFunction_('runDailyFilterCountsStart');
}

/** Permanent local entrypoint for the currently configured remote batch task. */
function runLotteonRemoteTaskStart() {
  return runLotteonRemoteTaskFunction_('runLotteonRemoteTaskStartRemote_');
}

/** Permanent local continuation handler for the currently configured remote batch task. */
function runLotteonRemoteTaskContinue() {
  return runLotteonRemoteTaskFunction_('runLotteonRemoteTaskContinueRemote_');
}

/**
 * Run once after installing v1.17.
 * This explicit local MailApp reference causes Apps Script to request the
 * script.send_mail scope, sends a test notice, and proves future automatic
 * terminal notices can be delivered.
 */
function authorizeLotteonRemoteNotifications() {
  const quota = MailApp.getRemainingDailyQuota();
  MailApp.sendEmail(
    LOTTEON_REMOTE_NOTICE_EMAIL,
    '[LOTTEON 자동작업 결과][TEST] 완료 알림 권한 승인',
    'LOTTEON 자동작업 완료 알림 권한 승인이 정상적으로 완료되었습니다.\n\n' +
    '브리지 버전: ' + LOTTEON_TRIGGER_BRIDGE_VERSION + '\n' +
    '남은 일일 수신자 한도: ' + quota
  );
  try {
    SpreadsheetApp.getUi().alert(
      'LOTTEON 완료 알림 권한 승인 완료\n\n테스트 메일을 발송했습니다.\n브리지 버전: ' +
      LOTTEON_TRIGGER_BRIDGE_VERSION
    );
  } catch (ignore) {}
  return {ok: true, version: LOTTEON_TRIGGER_BRIDGE_VERSION, quota: quota};
}

function runLotteonRemoteTaskFunction_(entryName) {
  const safeName = String(entryName || '').trim();
  const allowed = {
    runLotteonRemoteTaskStartRemote_: true,
    runLotteonRemoteTaskContinueRemote_: true
  };
  if (!allowed[safeName]) throw new Error('허용되지 않은 원격 작업 진입점: ' + safeName);

  const response = UrlFetchApp.fetch(LOTTEON_REMOTE_TASK_URL + '?ts=' + new Date().getTime(), {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true
  });
  const status = response.getResponseCode();
  const code = response.getContentText('UTF-8');
  if (status < 200 || status >= 300) {
    throw new Error('원격 작업 로드 실패 HTTP ' + status + '\n' + LOTTEON_REMOTE_TASK_URL + '\n' + code.slice(0, 500));
  }

  return eval(code + '\n\n; if (typeof ' + safeName + ' !== "function") throw new Error("원격 작업 진입점 없음: ' + safeName + '"); ' + safeName + '();');
}
