/**
 * Current remote batch task: PR38 v6.70 production apply.
 * Maintained on the permanent `remote/task-runner` branch.
 */
const LOTTEON_REMOTE_TASK_VERSION = 'PR38-v1.0';
const LOTTEON_REMOTE_MAIN_CODE_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Code.gs';
const LOTTEON_REMOTE_MAIN_PATCH_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_24_bootstrap_auto_continue.gs';
const LOTTEON_REMOTE_TASK_SOURCE_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/codex/issue-38-pr33-production-apply/PR38_Remote_Production_Apply.gs';

function runLotteonRemoteTaskStartRemote_() {
  return runLotteonRemotePr38_('runPr38ProductionApplyStart');
}

function runLotteonRemoteTaskContinueRemote_() {
  return runLotteonRemotePr38_('runPr38ProductionApplyContinue');
}

function runLotteonRemotePr38_(entryName) {
  const code = fetchRemoteTaskText_(LOTTEON_REMOTE_MAIN_CODE_URL, 'main Code.gs');
  const patch = fetchRemoteTaskText_(LOTTEON_REMOTE_MAIN_PATCH_URL, 'main patch bootstrap');
  const runner = fetchRemoteTaskText_(LOTTEON_REMOTE_TASK_SOURCE_URL, 'PR38 runner');
  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(entryName)) throw new Error('잘못된 원격 진입점: ' + entryName);
  const bundle = code + '\n\n;\n\n' + patch + '\n\n;\n\n' + runner;
  return eval(bundle + '\n\n; if (typeof ' + entryName + ' !== "function") throw new Error("PR38 원격 진입점 없음: ' + entryName + '"); ' + entryName + '();');
}

function fetchRemoteTaskText_(url, label) {
  const response = UrlFetchApp.fetch(url + '?ts=' + new Date().getTime(), {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true
  });
  const status = response.getResponseCode();
  const text = response.getContentText('UTF-8');
  if (status < 200 || status >= 300) {
    throw new Error(label + ' 로드 실패 HTTP ' + status + '\n' + url + '\n' + text.slice(0, 500));
  }
  return text;
}
