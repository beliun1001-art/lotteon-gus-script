/**
 * LOTTEON Google Sheets Apps Script GitHub Remote Loader v1.5
 *
 * Apps Script에는 이 파일만 붙여넣고,
 * 실제 운영 코드는 GitHub의 Code.gs + Patch_v6_01_daily_filter_auto.gs를 Raw URL로 불러와 실행합니다.
 *
 * v1.5:
 * - v1.4 메뉴 단순화 유지
 * - 고급/복구에서 K2 전용/과거 직접패치/초초경량 메뉴 숨김
 * - 실행 wrapper는 유지해 기존 트리거/과거 호출 호환성 보호
 */

const LOTTEON_GITHUB_CODE_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Code.gs';
const LOTTEON_GITHUB_PATCH_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_24_bootstrap_auto_continue.gs';
const LOTTEON_FILTER_COUNT_LIGHTWEIGHT_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/FilterCountLightweight_v6_56.gs';
const LOTTEON_GITHUB_README_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/README.md';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const mainMenu = ui.createMenu('LOTTEON 자동화');

  mainMenu
    .addItem('GitHub 코드 연결 테스트', 'testGitHubRemoteCode')
    .addSeparator()
    .addItem('① 변경사항 반영 실행', 'runPendingChangesApproval')
    .addItem('⑤ 대시보드만 빠른 갱신', 'refreshDashboardFastOnly')
    .addSeparator()
    .addItem('필터별_상품수 안전 갱신 시작/이어실행', 'runDailyFilterCountsOnceManual')
    .addItem('필터별_상품수 안전 갱신 상태 확인', 'showDailyFilterCountsStatus')
    .addItem('필터별_상품수 안전 갱신 초기화', 'resetDailyFilterCountsSafeState')
    .addSeparator()
    .addItem('③ 쿠팡재전송_로그 갱신', 'createRetransmitLogSheet')
    .addItem('④ 핵심요약+대시보드 갱신', 'refreshCoreSummaryAndDashboardWithRetransmitLogDates')
    .addItem('⑥ 시트 정리: 운영 시트만 표시', 'showOnlyMainSheets')
    .addSeparator();

  const settingsMenu = ui.createMenu('설정/관리')
    .addItem('GitHub 로더 권한 승인', 'authorizeLotteonLoader')
    .addSeparator()
    .addItem('변경감지 기능 시작', 'startChangeDetectionApproval')
    .addItem('변경감지 기능 중지', 'stopChangeDetectionApproval')
    .addSeparator()
    .addItem('필터별_상품수 안전 갱신 시작(매일 06:10)', 'startDailyFilterCountsSchedule')
    .addItem('필터별_상품수 안전 갱신 중지', 'stopDailyFilterCountsSchedule')
    .addSeparator()
    .addItem('API 인증값 저장', 'saveApiCredentials')
    .addItem('API 연결 테스트', 'testLotteonApiConnection')
    .addSeparator()
    .addItem('시트 복구: 전체 시트 표시', 'showAllSheets');

  const advancedMenu = ui.createMenu('고급/복구')
    .addItem('전체 검수 리포트 생성', 'generateAuditReport')
    .addItem('변경감지 상태 확인', 'showChangeDetectionStatus')
    .addItem('변경감지 플래그 초기화', 'resetChangeDetectionFlags')
    .addSeparator()
    .addItem('필터별_상품수 수동 1페이지 실행', 'refreshFilterCountsFast')
    .addItem('필터별_상품수 이어실행 초기화', 'resetFilterListResumeProgress')
    .addItem('필터_대시보드 갱신', 'refreshFilterDashboardFastOnly')
    .addSeparator()
    .addItem('자동/예약 트리거 전체 정리', 'cleanupAllAutoRefreshTriggers');

  mainMenu
    .addSubMenu(settingsMenu)
    .addSubMenu(advancedMenu)
    .addToUi();
}

function authorizeLotteonLoader() {
  const started = new Date();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  const authInfo = {
    spreadsheetName: ss.getName(),
    spreadsheetId: ss.getId(),
    sheetName: sheet ? sheet.getName() : '',
    startedAt: started.toISOString()
  };

  PropertiesService.getScriptProperties().setProperty('LOTTEON_LOADER_AUTH_AT', authInfo.startedAt);
  PropertiesService.getScriptProperties().setProperty('LOTTEON_LOADER_AUTH_VERSION', 'v1.5');
  CacheService.getScriptCache().put('LOTTEON_LOADER_AUTH_TEST', 'OK', 30);

  const lock = LockService.getScriptLock();
  try {
    lock.tryLock(1000);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }

  ScriptApp.getProjectTriggers();
  let tempTriggerCreated = 'N';
  try {
    const tempTrigger = ScriptApp.newTrigger('__lotteonLoaderAuthNoop')
      .timeBased()
      .after(60 * 60 * 1000)
      .create();
    tempTriggerCreated = 'Y';
    ScriptApp.deleteTrigger(tempTrigger);
  } catch (e) {
    Logger.log('임시 트리거 생성/삭제 중 경고: ' + e);
  }

  const readmeResponse = UrlFetchApp.fetch(LOTTEON_GITHUB_README_URL + '?auth_test=' + new Date().getTime(), {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true
  });
  const readmeCode = String(readmeResponse.getResponseCode());
  if (Number(readmeCode) < 200 || Number(readmeCode) >= 300) {
    throw new Error('README 연결 실패 HTTP ' + readmeCode + ': ' + readmeResponse.getContentText().slice(0, 300));
  }

  const msg =
    'GitHub 로더 권한 승인 완료' +
    ' / loader=v1.5' +
    ' / spreadsheet=' + authInfo.spreadsheetName +
    ' / sheet=' + authInfo.sheetName +
    ' / readmeHTTP=' + readmeCode +
    ' / tempTrigger=' + tempTriggerCreated;

  PropertiesService.getScriptProperties().setProperty('LOTTEON_LOADER_AUTH_RESULT', msg);
  Logger.log(msg);

  try {
    ss.toast('GitHub 로더 권한 승인 완료 v1.5', 'LOTTEON 자동화', 5);
  } catch (e) {}

  return msg;
}

function __lotteonLoaderAuthNoop() {}

function fetchGitHubRemoteCode_() {
  const ts = new Date().getTime();
  const response = UrlFetchApp.fetch(LOTTEON_GITHUB_CODE_URL + '?ts=' + ts, {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true
  });
  const code = response.getResponseCode();
  const text = response.getContentText('UTF-8');

  if (code < 200 || code >= 300) {
    throw new Error('GitHub Code.gs 로드 실패 HTTP ' + code + ': ' + text.slice(0, 500));
  }
  if (!text || text.indexOf('function runPendingChangesApproval') < 0) {
    throw new Error('GitHub Code.gs 내용이 올바르지 않습니다. runPendingChangesApproval 함수를 찾지 못했습니다.');
  }

  const patchResponse = UrlFetchApp.fetch(LOTTEON_GITHUB_PATCH_URL + '?ts=' + ts, {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true
  });
  const patchCode = patchResponse.getResponseCode();
  const patchText = patchResponse.getContentText('UTF-8');

  if (patchCode < 200 || patchCode >= 300) {
    throw new Error('GitHub Patch 로드 실패 HTTP ' + patchCode + ': ' + patchText.slice(0, 500));
  }
  if (!patchText || patchText.indexOf('function startDailyFilterCountsSchedule') < 0) {
    throw new Error('GitHub Patch 내용이 올바르지 않습니다. startDailyFilterCountsSchedule marker를 찾지 못했습니다.');
  }

  return text + '\n\n' + patchText;
}

function runRemoteFunction_(functionName, args) {
  args = args || [];
  const code = fetchGitHubRemoteCode_();
  const safeName = String(functionName || '').trim();
  if (!/^[A-Za-z_$][A-Za-z0-9_$]*$/.test(safeName)) {
    throw new Error('허용되지 않는 함수명: ' + safeName);
  }
  return eval(code + '\n' + safeName + '.apply(null, args);');
}

function runFilterCountLightweightFunction_(functionName) {
  const started = new Date().getTime();
  const response = UrlFetchApp.fetch(LOTTEON_FILTER_COUNT_LIGHTWEIGHT_URL + '?ts=' + started, { method: 'get', muteHttpExceptions: true, followRedirects: true });
  const text = response.getContentText('UTF-8');
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) throw new Error('FilterCountLightweight load failed HTTP ' + response.getResponseCode());
  PropertiesService.getScriptProperties().setProperty('LOTTEON_FILTER_COUNT_LIGHTWEIGHT_LOAD_METRICS', JSON.stringify({ remoteFetchFiles: 1, runnerEntryMs: new Date().getTime() - started, at: new Date().toISOString() }));
  return eval(text + '\n' + functionName + '();');
}

function testGitHubRemoteCode() {
  const code = fetchGitHubRemoteCode_();
  const versionMatch = code.match(/LOTTEON_PATCH_BOOTSTRAP_VERSION\s*=\s*['\"]([^'\"]+)['\"]/) || code.match(/v\d+\.\d+(?:\.\d+)?[^\n]*/i);
  SpreadsheetApp.getUi().alert(
    'GitHub 코드 연결 성공\n\n' +
    'Code.gs Raw URL:\n' + LOTTEON_GITHUB_CODE_URL + '\n\n' +
    'Patch Raw URL:\n' + LOTTEON_GITHUB_PATCH_URL + '\n\n' +
    '로드 크기: ' + code.length.toLocaleString() + '자\n' +
    '버전 추정: ' + (versionMatch ? (versionMatch[1] || versionMatch[0]) : '확인 필요') + '\n' +
    '현재 구조: Code.gs + v6.05 patch bootstrap 포함'
  );
}

// 상단 메뉴 wrapper
function runPendingChangesApproval() { return runRemoteFunction_('runPendingChangesApproval'); }
function refreshFilterCountsFast() { return runRemoteFunction_('refreshFilterCountsFast'); }
function createRetransmitLogSheet() { return runRemoteFunction_('createRetransmitLogSheet'); }
function refreshCoreSummaryAndDashboardWithRetransmitLogDates() { return runRemoteFunction_('refreshCoreSummaryAndDashboardWithRetransmitLogDates'); }
function refreshDashboardFastOnly() { return runRemoteFunction_('refreshDashboardFastOnly'); }
function showOnlyMainSheets() { return runRemoteFunction_('showOnlyMainSheets'); }

// 점검/검수 wrapper
function generateAuditReport() { return runRemoteFunction_('generateAuditReport'); }
function showStableAutoRefreshStatus() { return runRemoteFunction_('showStableAutoRefreshStatus'); }
function showChangeDetectionStatus() { return runRemoteFunction_('showChangeDetectionStatus'); }
function refreshManualApiCheckOnly() { return runRemoteFunction_('refreshManualApiCheckOnly'); }
function refreshFilterDashboardFastOnly() { return runRemoteFunction_('refreshFilterDashboardFastOnly'); }

// 설정/초기화 wrapper
function startChangeDetectionApproval() { return runRemoteFunction_('startChangeDetectionApproval'); }
function stopChangeDetectionApproval() { return runRemoteFunction_('stopChangeDetectionApproval'); }
function resetChangeDetectionFlags() { return runRemoteFunction_('resetChangeDetectionFlags'); }
function resetFilterListResumeProgress() { return runRemoteFunction_('resetFilterListResumeProgress'); }
function resetDailyFilterCountsSafeState() { return runFilterCountLightweightFunction_('resetDailyFilterCountsSafeState'); }
function startDailyFilterCountsSchedule() { return runFilterCountLightweightFunction_('startDailyFilterCountsSchedule'); }
function stopDailyFilterCountsSchedule() { return runFilterCountLightweightFunction_('stopDailyFilterCountsSchedule'); }
function runDailyFilterCountsOnceManual() { return runFilterCountLightweightFunction_('runDailyFilterCountsOnceManual'); }
function showDailyFilterCountsStatus() { return runFilterCountLightweightFunction_('showDailyFilterCountsStatus'); }
function runDailyFilterCountsStart() { return runFilterCountLightweightFunction_('runDailyFilterCountsStart'); }
function runDailyFilterCountsContinue() { return runFilterCountLightweightFunction_('runDailyFilterCountsContinue'); }
function cleanupAllAutoRefreshTriggers() { return runRemoteFunction_('cleanupAllAutoRefreshTriggers'); }
function saveApiCredentials() { return runRemoteFunction_('saveApiCredentials'); }
function testLotteonApiConnection() { return runRemoteFunction_('testLotteonApiConnection'); }
function showAllSheets() { return runRemoteFunction_('showAllSheets'); }

// 과거/복구용 wrapper는 메뉴에서 숨겼지만 기존 트리거/수동 호출 호환성을 위해 유지
function diagnoseRetransmitLogFilterDateK2() { return runRemoteFunction_('diagnoseRetransmitLogFilterDateK2'); }
function patchK2RetransmitLogDateFromApi() { return runRemoteFunction_('patchK2RetransmitLogDateFromApi'); }
function patchRetransmitLogDatesFromFilterSummary() { return runRemoteFunction_('patchRetransmitLogDatesFromFilterSummary'); }
function cleanFutureDatesInCoupangWorkLog() { return runRemoteFunction_('cleanFutureDatesInCoupangWorkLog'); }
function runStableAutoRefreshOnce() { return runRemoteFunction_('runStableAutoRefreshOnce'); }

// 설치형 onEdit 트리거 wrapper
function handleWatchedSheetEdit(e) { return runRemoteFunction_('handleWatchedSheetEdit', [e]); }

