/**
 * LOTTEON Apps Script permanent trigger bridges v1.16
 *
 * One-time local install. Future temporary batch runners are updated only in
 * GitHub's permanent `remote/task-runner/RemoteTaskCurrent.gs` slot.
 */
const LOTTEON_TRIGGER_BRIDGE_VERSION = 'v1.16';
const LOTTEON_REMOTE_TASK_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/remote/task-runner/RemoteTaskCurrent.gs';

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
