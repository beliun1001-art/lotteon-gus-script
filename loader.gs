/**
 * LOTTEON Google Sheets Apps Script GitHub Remote Loader v1.5
 *
 * Apps Script?먮뒗 ???뚯씪留?遺숈뿬?ｊ퀬,
 * ?ㅼ젣 ?댁쁺 肄붾뱶??GitHub??Code.gs + Patch_v6_01_daily_filter_auto.gs瑜?Raw URL濡?遺덈윭? ?ㅽ뻾?⑸땲??
 *
 * v1.5:
 * - v1.4 硫붾돱 ?⑥닚???좎?
 * - 怨좉툒/蹂듦뎄?먯꽌 K2 ?꾩슜/怨쇨굅 吏곸젒?⑥튂/珥덉큹寃쎈웾 硫붾돱 ?④?
 * - ?ㅽ뻾 wrapper???좎???湲곗〈 ?몃━嫄?怨쇨굅 ?몄텧 ?명솚??蹂댄샇
 */

const LOTTEON_GITHUB_CODE_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Code.gs';
const LOTTEON_GITHUB_PATCH_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_24_bootstrap_auto_continue.gs';
const LOTTEON_GITHUB_README_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/README.md';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const mainMenu = ui.createMenu('LOTTEON ?먮룞??);

  mainMenu
    .addItem('GitHub 肄붾뱶 ?곌껐 ?뚯뒪??, 'testGitHubRemoteCode')
    .addSeparator()
    .addItem('??蹂寃쎌궗??諛섏쁺 ?ㅽ뻾', 'runPendingChangesApproval')
    .addItem('????쒕낫?쒕쭔 鍮좊Ⅸ 媛깆떊', 'refreshDashboardFastOnly')
    .addSeparator()
    .addItem('?꾪꽣蹂??곹뭹???덉쟾 媛깆떊 ?쒖옉/?댁뼱?ㅽ뻾', 'runDailyFilterCountsOnceManual')
    .addItem('?꾪꽣蹂??곹뭹???덉쟾 媛깆떊 ?곹깭 ?뺤씤', 'showDailyFilterCountsStatus')
    .addItem('?꾪꽣蹂??곹뭹???덉쟾 媛깆떊 珥덇린??, 'resetDailyFilterCountsSafeState')
    .addSeparator()
    .addItem('??荑좏뙜?ъ쟾??濡쒓렇 媛깆떊', 'createRetransmitLogSheet')
    .addItem('???듭떖?붿빟+??쒕낫??媛깆떊', 'refreshCoreSummaryAndDashboardWithRetransmitLogDates')
    .addItem('???쒗듃 ?뺣━: ?댁쁺 ?쒗듃留??쒖떆', 'showOnlyMainSheets')
    .addSeparator();

  const settingsMenu = ui.createMenu('?ㅼ젙/愿由?)
    .addItem('GitHub 濡쒕뜑 沅뚰븳 ?뱀씤', 'authorizeLotteonLoader')
    .addSeparator()
    .addItem('蹂寃쎄컧吏 湲곕뒫 ?쒖옉', 'startChangeDetectionApproval')
    .addItem('蹂寃쎄컧吏 湲곕뒫 以묒?', 'stopChangeDetectionApproval')
    .addSeparator()
    .addItem('?꾪꽣蹂??곹뭹???덉쟾 媛깆떊 ?쒖옉(留ㅼ씪 06:10)', 'startDailyFilterCountsSchedule')
    .addItem('?꾪꽣蹂??곹뭹???덉쟾 媛깆떊 以묒?', 'stopDailyFilterCountsSchedule')
    .addSeparator()
    .addItem('API ?몄쬆媛????, 'saveApiCredentials')
    .addItem('API ?곌껐 ?뚯뒪??, 'testLotteonApiConnection')
    .addSeparator()
    .addItem('?쒗듃 蹂듦뎄: ?꾩껜 ?쒗듃 ?쒖떆', 'showAllSheets');

  const advancedMenu = ui.createMenu('怨좉툒/蹂듦뎄')
    .addItem('?꾩껜 寃??由ы룷???앹꽦', 'generateAuditReport')
    .addItem('蹂寃쎄컧吏 ?곹깭 ?뺤씤', 'showChangeDetectionStatus')
    .addItem('蹂寃쎄컧吏 ?뚮옒洹?珥덇린??, 'resetChangeDetectionFlags')
    .addSeparator()
    .addItem('?꾪꽣蹂??곹뭹???섎룞 1?섏씠吏 ?ㅽ뻾', 'refreshFilterCountsFast')
    .addItem('?꾪꽣蹂??곹뭹???댁뼱?ㅽ뻾 珥덇린??, 'resetFilterListResumeProgress')
    .addItem('?꾪꽣_??쒕낫??媛깆떊', 'refreshFilterDashboardFastOnly')
    .addSeparator()
    .addItem('?먮룞/?덉빟 ?몃━嫄??꾩껜 ?뺣━', 'cleanupAllAutoRefreshTriggers');

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
    Logger.log('?꾩떆 ?몃━嫄??앹꽦/??젣 以?寃쎄퀬: ' + e);
  }

  const readmeResponse = UrlFetchApp.fetch(LOTTEON_GITHUB_README_URL + '?auth_test=' + new Date().getTime(), {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true
  });
  const readmeCode = String(readmeResponse.getResponseCode());
  if (Number(readmeCode) < 200 || Number(readmeCode) >= 300) {
    throw new Error('README ?곌껐 ?ㅽ뙣 HTTP ' + readmeCode + ': ' + readmeResponse.getContentText().slice(0, 300));
  }

  const msg =
    'GitHub 濡쒕뜑 沅뚰븳 ?뱀씤 ?꾨즺' +
    ' / loader=v1.5' +
    ' / spreadsheet=' + authInfo.spreadsheetName +
    ' / sheet=' + authInfo.sheetName +
    ' / readmeHTTP=' + readmeCode +
    ' / tempTrigger=' + tempTriggerCreated;

  PropertiesService.getScriptProperties().setProperty('LOTTEON_LOADER_AUTH_RESULT', msg);
  Logger.log(msg);

  try {
    ss.toast('GitHub 濡쒕뜑 沅뚰븳 ?뱀씤 ?꾨즺 v1.5', 'LOTTEON ?먮룞??, 5);
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
    throw new Error('GitHub Code.gs 濡쒕뱶 ?ㅽ뙣 HTTP ' + code + ': ' + text.slice(0, 500));
  }
  if (!text || text.indexOf('function runPendingChangesApproval') < 0) {
    throw new Error('GitHub Code.gs ?댁슜???щ컮瑜댁? ?딆뒿?덈떎. runPendingChangesApproval ?⑥닔瑜?李얠? 紐삵뻽?듬땲??');
  }

  const patchResponse = UrlFetchApp.fetch(LOTTEON_GITHUB_PATCH_URL + '?ts=' + ts, {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true
  });
  const patchCode = patchResponse.getResponseCode();
  const patchText = patchResponse.getContentText('UTF-8');

  if (patchCode < 200 || patchCode >= 300) {
    throw new Error('GitHub Patch 濡쒕뱶 ?ㅽ뙣 HTTP ' + patchCode + ': ' + patchText.slice(0, 500));
  }
  if (!patchText || patchText.indexOf('function startDailyFilterCountsSchedule') < 0) {
    throw new Error('GitHub Patch ?댁슜???щ컮瑜댁? ?딆뒿?덈떎. startDailyFilterCountsSchedule marker瑜?李얠? 紐삵뻽?듬땲??');
  }

  return text + '\n\n' + patchText;
}

function runRemoteFunction_(functionName, args) {
  args = args || [];
  const code = fetchGitHubRemoteCode_();
  const safeName = String(functionName || '').trim();
  if (!/^[A-Za-z_$][A-Za-z0-9_$]*$/.test(safeName)) {
    throw new Error('?덉슜?섏? ?딅뒗 ?⑥닔紐? ' + safeName);
  }
  return eval(code + '\n' + safeName + '.apply(null, args);');
}

function testGitHubRemoteCode() {
  const code = fetchGitHubRemoteCode_();
  const versionMatch = code.match(/LOTTEON_PATCH_BOOTSTRAP_VERSION\s*=\s*['\"]([^'\"]+)['\"]/) || code.match(/v\d+\.\d+(?:\.\d+)?[^\n]*/i);
  SpreadsheetApp.getUi().alert(
    'GitHub 肄붾뱶 ?곌껐 ?깃났\n\n' +
    'Code.gs Raw URL:\n' + LOTTEON_GITHUB_CODE_URL + '\n\n' +
    'Patch Raw URL:\n' + LOTTEON_GITHUB_PATCH_URL + '\n\n' +
    '濡쒕뱶 ?ш린: ' + code.length.toLocaleString() + '??n' +
    '踰꾩쟾 異붿젙: ' + (versionMatch ? (versionMatch[1] || versionMatch[0]) : '?뺤씤 ?꾩슂') + '\n' +
    '?꾩옱 援ъ“: Code.gs + v6.05 patch bootstrap ?ы븿'
  );
}

// ?곷떒 硫붾돱 wrapper
function runPendingChangesApproval() { return runRemoteFunction_('runPendingChangesApproval'); }
function refreshFilterCountsFast() { return runRemoteFunction_('refreshFilterCountsFast'); }
function createRetransmitLogSheet() { return runRemoteFunction_('createRetransmitLogSheet'); }
function refreshCoreSummaryAndDashboardWithRetransmitLogDates() { return runRemoteFunction_('refreshCoreSummaryAndDashboardWithRetransmitLogDates'); }
function refreshDashboardFastOnly() { return runRemoteFunction_('refreshDashboardFastOnly'); }
function showOnlyMainSheets() { return runRemoteFunction_('showOnlyMainSheets'); }

// ?먭?/寃??wrapper
function generateAuditReport() { return runRemoteFunction_('generateAuditReport'); }
function showStableAutoRefreshStatus() { return runRemoteFunction_('showStableAutoRefreshStatus'); }
function showChangeDetectionStatus() { return runRemoteFunction_('showChangeDetectionStatus'); }
function refreshManualApiCheckOnly() { return runRemoteFunction_('refreshManualApiCheckOnly'); }
function refreshFilterDashboardFastOnly() { return runRemoteFunction_('refreshFilterDashboardFastOnly'); }

// ?ㅼ젙/珥덇린??wrapper
function startChangeDetectionApproval() { return runRemoteFunction_('startChangeDetectionApproval'); }
function stopChangeDetectionApproval() { return runRemoteFunction_('stopChangeDetectionApproval'); }
function resetChangeDetectionFlags() { return runRemoteFunction_('resetChangeDetectionFlags'); }
function resetFilterListResumeProgress() { return runRemoteFunction_('resetFilterListResumeProgress'); }
function resetDailyFilterCountsSafeState() { return runRemoteFunction_('resetDailyFilterCountsSafeState'); }
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

// 怨쇨굅/蹂듦뎄??wrapper??硫붾돱?먯꽌 ?④꼈吏留?湲곗〈 ?몃━嫄??섎룞 ?몄텧 ?명솚?깆쓣 ?꾪빐 ?좎?
function diagnoseRetransmitLogFilterDateK2() { return runRemoteFunction_('diagnoseRetransmitLogFilterDateK2'); }
function patchK2RetransmitLogDateFromApi() { return runRemoteFunction_('patchK2RetransmitLogDateFromApi'); }
function patchRetransmitLogDatesFromFilterSummary() { return runRemoteFunction_('patchRetransmitLogDatesFromFilterSummary'); }
function cleanFutureDatesInCoupangWorkLog() { return runRemoteFunction_('cleanFutureDatesInCoupangWorkLog'); }
function runStableAutoRefreshOnce() { return runRemoteFunction_('runStableAutoRefreshOnce'); }

// ?ㅼ튂??onEdit ?몃━嫄?wrapper
function handleWatchedSheetEdit(e) { return runRemoteFunction_('handleWatchedSheetEdit', [e]); }

