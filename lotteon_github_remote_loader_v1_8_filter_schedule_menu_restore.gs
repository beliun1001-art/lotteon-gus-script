/**
 * LOTTEON Google Sheets Apps Script GitHub Remote Loader v1.10
 *
 * v1.10 핵심:
 * - v1.8 ScriptApp 권한 포함 구조 유지
 * - LOTTEON 자동화 / LOTTEON 서식 메뉴 유지
 * - 누락된 필터별_상품수 매일 자동 갱신 시작/중지 메뉴 복구
 * - GitHub patch bootstrap은 Patch_v6_24_bootstrap_auto_continue.gs 사용
 * - v6.53 이상 GitHub patch를 자동 로드
 * - LOTTEON 서식 메뉴에 금액/백분율 전용 표시 형식 메뉴 유지
 * - 운영 핵심 10개 시트만 표시/불필요 시트 삭제 메뉴명 반영
 *
 * 사용 방법:
 * - Apps Script의 기존 loader 전체를 이 파일 전체로 교체합니다.
 * - 저장 후 시트를 새로고침합니다.
 * - LOTTEON 자동화 → 설정/관리 → GitHub 로더 권한 승인을 한 번 실행합니다.
 */

const LOTTEON_GITHUB_CODE_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Code.gs';
const LOTTEON_GITHUB_PATCH_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_24_bootstrap_auto_continue.gs';
const LOTTEON_LOADER_VERSION = 'v1.10';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('LOTTEON 자동화')
    .addItem('GitHub 코드 연결 테스트', 'testLotteonGitHubConnection')
    .addSeparator()
    .addItem('① 변경사항 반영 실행', 'runPendingChangesApproval')
    .addItem('⑤ 대시보드만 빠른 갱신', 'refreshDashboardFastOnly')
    .addSeparator()
    .addItem('필터별_상품수 자동 갱신 지금 시작', 'runDailyFilterCountsOnceManual')
    .addItem('필터별_상품수 자동 상태 확인', 'showDailyFilterCountsStatus')
    .addItem('필터별_상품수 매일 자동 갱신 시작(06:10)', 'startDailyFilterCountsSchedule')
    .addItem('필터별_상품수 매일 자동 갱신 중지', 'stopDailyFilterCountsSchedule')
    .addSeparator()
    .addItem('③ 쿠팡재전송_로그 갱신', 'createRetransmitLogSheet')
    .addItem('④ 핵심요약+대시보드 갱신', 'refreshCoreSummaryAndDashboardWithRetransmitLogDates')
    .addItem('⑥ 운영 핵심 10개만 표시/불필요 삭제', 'showOperationSheetsOnly')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('설정/관리')
        .addItem('GitHub 로더 권한 승인', 'authorizeLotteonLoader')
        .addItem('ScriptApp 트리거 권한 확인', 'authorizeLotteonScriptAppScope')
        .addItem('GitHub 코드 캐시 초기화', 'clearLotteonGitHubCodeCache')
        .addItem('필터별_상품수 매일 자동 갱신 시작(06:10)', 'startDailyFilterCountsSchedule')
        .addItem('필터별_상품수 매일 자동 갱신 중지', 'stopDailyFilterCountsSchedule')
        .addItem('필터별_상품수 이어실행 초기화', 'resetFilterListResumeState')
        .addSeparator()
        .addItem('변경감지 기능 시작', 'startChangeDetection')
        .addItem('변경감지 기능 중지', 'stopChangeDetection')
        .addItem('변경감지 플래그 초기화', 'resetChangeDetectionFlags')
    )
    .addSubMenu(
      ui.createMenu('고급/복구')
        .addItem('부가세 자료 생성/이어실행', 'generateVatReportsFullSeparated_v622')
        .addItem('운영 핵심 10개만 표시/불필요 삭제', 'showOperationSheetsOnly')
        .addItem('시트 복구: 전체 시트 표시', 'showAllSheets')
        .addItem('API 인증값 저장', 'saveApiCredentials')
        .addItem('API 연결 테스트', 'testApiConnection')
    )
    .addToUi();

  ui.createMenu('LOTTEON 서식')
    .addItem('열너비 안전조정 시작/이어실행', 'runColumnWidthAutoAdjustStep_v623')
    .addItem('열너비 안전조정 상태 확인', 'showColumnWidthAutoAdjustStatus_v623')
    .addItem('열너비 안전조정 초기화', 'resetColumnWidthAutoAdjust_v623')
    .addSeparator()
    .addItem('금액 1,000단위 쉼표 적용', 'applyAmountThousandsFormat_v649')
    .addItem('백분율 % 적용', 'applyPercentOneDecimalFormat_v649')
    .addSeparator()
    .addItem('표시서식만 빠른 정리', 'applyDisplayStandardsOnlyFast_v623')
    .addToUi();
}

function authorizeLotteonLoader() {
  SpreadsheetApp.getActive();
  UrlFetchApp.fetch(LOTTEON_GITHUB_CODE_URL, { muteHttpExceptions: true, followRedirects: true });
  UrlFetchApp.fetch(LOTTEON_GITHUB_PATCH_URL, { muteHttpExceptions: true, followRedirects: true });
  authorizeLotteonScriptAppScope();
  SpreadsheetApp.getUi().alert('GitHub 로더 권한 승인 완료\n\nScriptApp 트리거 권한까지 확인했습니다.\n스프레드시트를 새로고침한 뒤 메뉴를 다시 실행하세요.');
}

function authorizeLotteonScriptAppScope() {
  const triggers = ScriptApp.getProjectTriggers();
  const handlerName = 'lotteonDummyTriggerScopeHandler_';
  let dummy = null;
  try {
    dummy = ScriptApp.newTrigger(handlerName).timeBased().after(60 * 60 * 1000).create();
  } catch (e) {}
  if (dummy) {
    try { ScriptApp.deleteTrigger(dummy); } catch (e) {}
  }
  try {
    SpreadsheetApp.getUi().alert('ScriptApp 트리거 권한 확인 완료\n\n현재 프로젝트 트리거 수: ' + triggers.length);
  } catch (e) {}
}

function lotteonDummyTriggerScopeHandler_() {}

function testLotteonGitHubConnection() {
  const bundle = loadLotteonRemoteBundle_();
  const version = detectRemoteVersion_(bundle);
  SpreadsheetApp.getUi().alert(
    'GitHub 코드 연결 성공\n\n' +
    'Code.gs Raw URL:\n' + LOTTEON_GITHUB_CODE_URL + '\n\n' +
    'Patch Raw URL:\n' + LOTTEON_GITHUB_PATCH_URL + '\n\n' +
    '로더 버전: ' + LOTTEON_LOADER_VERSION + '\n' +
    '로드 크기: ' + bundle.length.toLocaleString('ko-KR') + '자\n' +
    '버전 추정: ' + version + '\n' +
    '캐시 방식: 사용 안 함\n' +
    'ScriptApp 권한: 로컬 loader에 포함\n' +
    '필터 자동예약 메뉴: 복구됨\n\n' +
    '필터별_상품수 매일 자동 갱신 시작/중지 메뉴를 사용할 수 있습니다.'
  );
}

function clearLotteonGitHubCodeCache() {
  ['LOTTEON_REMOTE_CODE_BUNDLE_V14','LOTTEON_REMOTE_CODE_BUNDLE_V13','LOTTEON_REMOTE_CODE_BUNDLE'].forEach(function(k){ try { CacheService.getScriptCache().remove(k); } catch(e) {} });
  PropertiesService.getScriptProperties().deleteProperty('LOTTEON_REMOTE_LAST_VERSION');
  SpreadsheetApp.getUi().alert('GitHub 코드 캐시를 초기화했습니다.\n\nv1.10은 대용량 코드를 캐시에 저장하지 않습니다.');
}

function loadLotteonRemoteBundle_() {
  const codeText = fetchTextOrThrow_(LOTTEON_GITHUB_CODE_URL, 'Code.gs');
  const patchText = fetchTextOrThrow_(LOTTEON_GITHUB_PATCH_URL, 'Patch bootstrap');
  const bundle = codeText + '\n\n;\n\n' + patchText;
  PropertiesService.getScriptProperties().setProperty('LOTTEON_REMOTE_LAST_VERSION', detectRemoteVersion_(bundle));
  return bundle;
}

function fetchTextOrThrow_(url, label) {
  const response = UrlFetchApp.fetch(url + '?ts=' + new Date().getTime(), { method: 'get', muteHttpExceptions: true, followRedirects: true });
  const code = response.getResponseCode();
  const text = response.getContentText('UTF-8');
  if (code < 200 || code >= 300) throw new Error(label + ' 로드 실패 HTTP ' + code + '\n' + url + '\n' + text.slice(0, 500));
  return text;
}

function runRemoteFunctionByName_(functionName) {
  const bundle = loadLotteonRemoteBundle_();
  const safeName = String(functionName || '').trim();
  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(safeName)) throw new Error('잘못된 함수명: ' + safeName);
  return eval(bundle + '\n\n; if (typeof ' + safeName + ' !== "function") { throw new Error("Remote function not found: ' + safeName + '"); }\n' + safeName + '();');
}

function runRemoteFirstAvailable_(functionNames) {
  const bundle = loadLotteonRemoteBundle_();
  const namesLiteral = JSON.stringify(functionNames || []);
  return eval(bundle + '\n\n; (function(){ var names = ' + namesLiteral + '; for (var i=0;i<names.length;i++){ try { if (typeof eval(names[i]) === "function") return eval(names[i])(); } catch(e) {} } throw new Error("Remote function not found: " + names.join(", ")); })();');
}

function detectRemoteVersion_(text) {
  const src = String(text || '');
  const m = src.match(/LOTTEON_PATCH_BOOTSTRAP_VERSION\s*=\s*['"]([^'"]+)['"]/);
  if (m) return m[1];
  const all = src.match(/v6\.\d+/g);
  return all && all.length ? all[all.length - 1] : '확인불가';
}

function runPendingChangesApproval() { return runRemoteFunctionByName_('runPendingChangesApproval'); }
function refreshDashboardFastOnly() { return runRemoteFunctionByName_('refreshDashboardFastOnly'); }
function runDailyFilterCountsOnceManual() { return runRemoteFirstAvailable_(['runDailyFilterCountsOnceManual', 'runDailyFilterCountsStep_', 'refreshFilterCountsFast']); }
function showDailyFilterCountsStatus() { return runRemoteFirstAvailable_(['showDailyFilterCountsStatus', 'showDailyFilterCountsStatus_v601', 'showFilterCountsStatus']); }
function startDailyFilterCountsSchedule() { return runRemoteFunctionByName_('startDailyFilterCountsSchedule'); }
function stopDailyFilterCountsSchedule() { return runRemoteFunctionByName_('stopDailyFilterCountsSchedule'); }
function createRetransmitLogSheet() { return runRemoteFunctionByName_('createRetransmitLogSheet'); }
function refreshCoreSummaryAndDashboardWithRetransmitLogDates() { return runRemoteFunctionByName_('refreshCoreSummaryAndDashboardWithRetransmitLogDates'); }
function showOperationSheetsOnly() { return runRemoteFirstAvailable_(['cleanupOperationSheets_v653', 'showOperationSheetsOnly', 'showOnlyOperationSheets', 'hideNonOperationSheets']); }
function showAllSheets() { return runRemoteFirstAvailable_(['showAllSheets', 'restoreAllSheetsVisible']); }
function saveApiCredentials() { return runRemoteFirstAvailable_(['saveApiCredentials', 'saveApiCredentialsMenu']); }
function testApiConnection() { return runRemoteFirstAvailable_(['testApiConnection', 'testApiConnectionMenu']); }
function startChangeDetection() { return runRemoteFirstAvailable_(['startChangeDetection', 'startChangeDetectionTrigger']); }
function stopChangeDetection() { return runRemoteFirstAvailable_(['stopChangeDetection', 'stopChangeDetectionTrigger']); }
function resetChangeDetectionFlags() { return runRemoteFirstAvailable_(['resetChangeDetectionFlags', 'resetChangeDetectionFlag']); }
function resetFilterListResumeState() { return runRemoteFirstAvailable_(['resetFilterListResumeState', 'resetFilterListResume']); }
function generateVatReportsFullSeparated_v622() { return runRemoteFunctionByName_('generateVatReportsFullSeparated_v622'); }
function runColumnWidthAutoAdjustStep_v623() { return runRemoteFunctionByName_('runColumnWidthAutoAdjustStep_v623'); }
function showColumnWidthAutoAdjustStatus_v623() { return runRemoteFunctionByName_('showColumnWidthAutoAdjustStatus_v623'); }
function resetColumnWidthAutoAdjust_v623() { return runRemoteFunctionByName_('resetColumnWidthAutoAdjust_v623'); }
function applyAmountThousandsFormat_v649() { return runRemoteFunctionByName_('applyAmountThousandsFormat_v649'); }
function applyPercentOneDecimalFormat_v649() { return runRemoteFunctionByName_('applyPercentOneDecimalFormat_v649'); }
function applyDisplayStandardsOnlyFast_v623() { return runRemoteFunctionByName_('applyDisplayStandardsOnlyFast_v623'); }
