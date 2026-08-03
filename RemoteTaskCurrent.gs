/**
 * Current remote batch task: PR33 v1.1 preview smoke.
 * Maintained on the permanent `remote/task-runner` branch.
 */
const LOTTEON_REMOTE_TASK_VERSION = 'PR33-v1.1';
const LOTTEON_REMOTE_TASK_SOURCE_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/codex/issue-32-vat-date-window-fallback/PR33_Preview_Smoke_v1_1.gs';

function runLotteonRemoteTaskStartRemote_() {
  return runLotteonRemotePr33_('runPr33PreviewSmoke');
}

function runLotteonRemoteTaskContinueRemote_() {
  return runLotteonRemotePr33_('runPr33PreviewSmokeContinue');
}

function runLotteonRemotePr33_(entryName) {
  const response = UrlFetchApp.fetch(LOTTEON_REMOTE_TASK_SOURCE_URL + '?ts=' + new Date().getTime(), {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true
  });
  const status = response.getResponseCode();
  let code = response.getContentText('UTF-8');
  if (status < 200 || status >= 300) {
    throw new Error('PR33 원격 runner 로드 실패 HTTP ' + status + '\n' + code.slice(0, 500));
  }

  // Continuation triggers must target a function that physically exists in Apps Script.
  code = code.replace(
    "const PR33_CONTINUE_HANDLER = 'runPr33PreviewSmokeContinue';",
    "const PR33_CONTINUE_HANDLER = 'runLotteonRemoteTaskContinue';"
  );

  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(entryName)) throw new Error('잘못된 원격 진입점: ' + entryName);
  return eval(code + '\n\n; if (typeof ' + entryName + ' !== "function") throw new Error("PR33 원격 진입점 없음: ' + entryName + '"); ' + entryName + '();');
}
