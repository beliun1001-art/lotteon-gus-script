/**
 * LOTTEON Google Sheets Apps Script GitHub Remote Loader v1.1
 *
 * Apps Script에는 이 파일만 붙여넣고,
 * 실제 운영 코드는 GitHub의 Code.gs를 Raw URL로 불러와 실행합니다.
 *
 * 주의:
 * - Code.gs가 약 400KB 이상이므로 Apps Script CacheService에는 저장하지 않습니다.
 * - GitHub의 Code.gs를 수정한 뒤에는 구글시트에서 메뉴를 다시 실행하면 최신 코드를 새로 불러옵니다.
 */

const LOTTEON_GITHUB_CODE_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Code.gs';

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
    .addItem('변경감지 기능 시작', 'startChangeDetectionApproval')
    .addItem('변경감지 기능 중지', 'stopChangeDetectionApproval')
    .addItem('변경감지 플래그 초기화', 'resetChangeDetectionFlags')
    .addItem('필터별_상품수 이어실행 초기화', 'resetFilterListResumeProgress')
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

function fetchGitHubRemoteCode_() {
  const response = UrlFetchApp.fetch(LOTTEON_GITHUB_CODE_URL + '?ts=' + new Date().getTime(), {
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

  return text;
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
  const versionMatch = code.match(/LOTTEON[^\n]*v\d+\.\d+[^\n]*/i) || code.match(/v\d+\.\d+[^\n]*/i);
  SpreadsheetApp.getUi().alert(
    'GitHub 코드 연결 성공\n\n' +
    'Raw URL:\n' + LOTTEON_GITHUB_CODE_URL + '\n\n' +
    '로드 크기: ' + code.length.toLocaleString() + '자\n' +
    '버전 추정: ' + (versionMatch ? versionMatch[0] : '확인 필요')
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
