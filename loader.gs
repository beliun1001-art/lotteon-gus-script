/**
 * LOTTEON Google Sheets Apps Script GitHub Remote Loader v1.2
 *
 * Apps Script에는 이 파일만 붙여넣고,
 * 실제 운영 코드는 GitHub의 Code.gs + Patch_v6_01_daily_filter_auto.gs를 Raw URL로 불러와 실행합니다.
 *
 * v1.2:
 * - 권한승인용 authorizeLotteonLoader 유지
 * - 필터별_상품수 매일 자동 갱신 메뉴/wrapper 추가
 * - Code.gs v6.00 뒤에 v6.01 patch 파일을 추가 로드
 */

const LOTTEON_GITHUB_CODE_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Code.gs';
const LOTTEON_GITHUB_PATCH_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_01_daily_filter_auto.gs';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const mainMenu = ui.createMenu('LOTTEON 자동화');

  mainMenu
    .addItem('① 변경사항 반영 실행', 'runPendingChangesApproval')
    .addItem('② 필터별_상품수 갱신/이어실행', 'refreshFilterCountsFast')
    .addItem('③ 쿠팡재전송_로그 갱신', 'createRetransmitLogSheet')
    .addItem('④ 핵심요약+대시보드 갱신', 'refreshCoreSummaryAndDashboardWithRetransmitLogDates')
    .addItem('⑤ 대시보드만 빠른 갱신', 'refreshDashboardFastOnly')
    .addItem('⑥ 시트 정리: 운영 시트만 표시', 'showOnlyMainSheets')
    .addSeparator();

  const checkMenu = ui.createMenu('점검/검수')
    .addItem('전체 검수 리포트 생성', 'generateAuditReport')
    .addItem('자동 갱신 상태 확인', 'showStableAutoRefreshStatus')
    .addItem('변경감지 상태 확인', 'showChangeDetectionStatus')
    .addItem('수동입력/API 검증 시트 갱신', 'refreshManualApiCheckOnly')
    .addItem('필터_대시보드 갱신', 'refreshFilterDashboardFastOnly');

  const setupMenu = ui.createMenu('설정/초기화')
    .addItem('GitHub 로더 권한 승인', 'authorizeLotteonLoader')
    .addSeparator()
    .addItem('변경감지 기능 시작', 'startChangeDetectionApproval')
    .addItem('변경감지 기능 중지', 'stopChangeDetectionApproval')
    .addItem('변경감지 플래그 초기화', 'resetChangeDetectionFlags')
    .addItem('필터별_상품수 이어실행 초기화', 'resetFilterListResumeProgress')
    .addSeparator()
    .addItem('필터별_상품수 자동 갱신 시작(매일 06:10)', 'startDailyFilterCountsSchedule')
    .addItem('필터별_상품수 자동 갱신 중지', 'stopDailyFilterCountsSchedule')
    .addItem('필터별_상품수 자동 갱신 지금 시작', 'runDailyFilterCountsOnceManual')
    .addItem('필터별_상품수 자동 상태 확인', 'showDailyFilterCountsStatus')
    .addSeparator()
    .addItem('자동/예약 트리거 전체 정리', 'cleanupAllAutoRefreshTriggers')
    .addSeparator()
    .addItem('API 인증값 저장', 'saveApiCredentials')
    .addItem('API 연결 테스트', 'testLotteonApiConnection')
    .addItem('시트 복구: 전체 시트 표시', 'showAllSheets')
    .addSeparator()
    .addItem('GitHub 코드 연결 테스트', 'testGitHubRemoteCode');

  const troubleMenu = ui.createMenu('문제해결')
    .addItem('K2 필터일자 진단(조회만)', 'diagnoseRetransmitLogFilterDateK2')
    .addItem('K2 날짜 직접반영(실제수정)', 'patchK2RetransmitLogDateFromApi')
    .addItem('쿠팡재전송_로그 날짜 직접패치', 'patchRetransmitLogDatesFromFilterSummary')
    .addItem('쿠팡재전송_로그 미래날짜 정리', 'cleanFutureDatesInCoupangWorkLog')
    .addItem('초초경량 자동 갱신 1회 실행', 'runStableAutoRefreshOnce');

  mainMenu
    .addSubMenu(checkMenu)
    .addSubMenu(setupMenu)
    .addSubMenu(troubleMenu)
    .addToUi();
}

function authorizeLotteonLoader() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  const ssId = ss.getId();
  const ssName = ss.getName();
  const sheetName = sheet ? sheet.getName() : '';

  PropertiesService.getScriptProperties().setProperty('LOTTEON_LOADER_AUTH_AT', new Date().toISOString());
  CacheService.getScriptCache().put('LOTTEON_LOADER_AUTH_TEST', 'OK', 30);

  const lock = LockService.getScriptLock();
  try {
    lock.tryLock(1000);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }

  ScriptApp.getProjectTriggers();

  const codeResponse = UrlFetchApp.fetch(LOTTEON_GITHUB_CODE_URL + '?auth_test=' + new Date().getTime(), {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true
  });
  const patchResponse = UrlFetchApp.fetch(LOTTEON_GITHUB_PATCH_URL + '?auth_test=' + new Date().getTime(), {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true
  });

  const code = codeResponse.getResponseCode();
  const patch = patchResponse.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('GitHub Code.gs 연결 실패 HTTP ' + code + ': ' + codeResponse.getContentText().slice(0, 500));
  }
  if (patch < 200 || patch >= 300) {
    throw new Error('GitHub Patch_v6_01 연결 실패 HTTP ' + patch + ': ' + patchResponse.getContentText().slice(0, 500));
  }

  SpreadsheetApp.getUi().alert(
    'GitHub 로더 권한 승인 완료\n\n' +
    '스프레드시트: ' + ssName + '\n' +
    '스프레드시트 ID: ' + ssId + '\n' +
    '현재 시트: ' + sheetName + '\n' +
    'Code.gs 연결: HTTP ' + code + '\n' +
    'Patch_v6_01 연결: HTTP ' + patch + '\n\n' +
    '이제 GitHub 코드 연결 테스트를 실행하세요.'
  );
}

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
    throw new Error('GitHub Patch_v6_01 로드 실패 HTTP ' + patchCode + ': ' + patchText.slice(0, 500));
  }
  if (!patchText || patchText.indexOf('function startDailyFilterCountsSchedule') < 0) {
    throw new Error('GitHub Patch_v6_01 내용이 올바르지 않습니다. startDailyFilterCountsSchedule 함수를 찾지 못했습니다.');
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

function testGitHubRemoteCode() {
  const code = fetchGitHubRemoteCode_();
  const versionMatch = code.match(/v\d+\.\d+[^\n]*/i);
  SpreadsheetApp.getUi().alert(
    'GitHub 코드 연결 성공\n\n' +
    'Code.gs Raw URL:\n' + LOTTEON_GITHUB_CODE_URL + '\n\n' +
    'Patch Raw URL:\n' + LOTTEON_GITHUB_PATCH_URL + '\n\n' +
    '로드 크기: ' + code.length.toLocaleString() + '자\n' +
    '버전 추정: ' + (versionMatch ? versionMatch[0] : '확인 필요') + '\n' +
    'v6.01 자동 필터 갱신 patch: 포함'
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
function startDailyFilterCountsSchedule() { return runRemoteFunction_('startDailyFilterCountsSchedule'); }
function stopDailyFilterCountsSchedule() { return runRemoteFunction_('stopDailyFilterCountsSchedule'); }
function runDailyFilterCountsOnceManual() { return runRemoteFunction_('runDailyFilterCountsOnceManual'); }
function showDailyFilterCountsStatus() { return runRemoteFunction_('showDailyFilterCountsStatus'); }
function runDailyFilterCountsStart() { return runRemoteFunction_('runDailyFilterCountsStart'); }
function runDailyFilterCountsContinue() { return runRemoteFunction_('runDailyFilterCountsContinue'); }
function cleanupAllAutoRefreshTriggers() { return runRemoteFunction_('cleanupAllAutoRefreshTriggers'); }
function saveApiCredentials() { return runRemoteFunction_('saveApiCredentials'); }
function testLotteonApiConnection() { return runRemoteFunction_('testLotteonApiConnection'); }
function showAllSheets() { return runRemoteFunction_('showAllSheets'); }

// 문제해결 wrapper
function diagnoseRetransmitLogFilterDateK2() { return runRemoteFunction_('diagnoseRetransmitLogFilterDateK2'); }
function patchK2RetransmitLogDateFromApi() { return runRemoteFunction_('patchK2RetransmitLogDateFromApi'); }
function patchRetransmitLogDatesFromFilterSummary() { return runRemoteFunction_('patchRetransmitLogDatesFromFilterSummary'); }
function cleanFutureDatesInCoupangWorkLog() { return runRemoteFunction_('cleanFutureDatesInCoupangWorkLog'); }
function runStableAutoRefreshOnce() { return runRemoteFunction_('runStableAutoRefreshOnce'); }

// 설치형 onEdit 트리거 wrapper
function handleWatchedSheetEdit(e) { return runRemoteFunction_('handleWatchedSheetEdit', [e]); }
