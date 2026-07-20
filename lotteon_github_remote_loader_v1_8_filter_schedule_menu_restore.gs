/**
 * LOTTEON Google Sheets Apps Script GitHub Remote Loader v1.11
 *
 * v1.10 ?듭떖:
 * - v1.8 ScriptApp 沅뚰븳 ?ы븿 援ъ“ ?좎?
 * - LOTTEON ?먮룞??/ LOTTEON ?쒖떇 硫붾돱 ?좎?
 * - ?꾨씫???꾪꽣蹂??곹뭹??留ㅼ씪 ?먮룞 媛깆떊 ?쒖옉/以묒? 硫붾돱 蹂듦뎄
 * - GitHub patch bootstrap? Patch_v6_24_bootstrap_auto_continue.gs ?ъ슜
 * - v6.53 ?댁긽 GitHub patch瑜??먮룞 濡쒕뱶
 * - LOTTEON ?쒖떇 硫붾돱??湲덉븸/諛깅텇???꾩슜 ?쒖떆 ?뺤떇 硫붾돱 ?좎?
 * - ?댁쁺 ?듭떖 10媛??쒗듃留??쒖떆/遺덊븘???쒗듃 ??젣 硫붾돱紐?諛섏쁺
 *
 * ?ъ슜 諛⑸쾿:
 * - Apps Script??湲곗〈 loader ?꾩껜瑜????뚯씪 ?꾩껜濡?援먯껜?⑸땲??
 * - ??????쒗듃瑜??덈줈怨좎묠?⑸땲??
 * - LOTTEON ?먮룞?????ㅼ젙/愿由???GitHub 濡쒕뜑 沅뚰븳 ?뱀씤????踰??ㅽ뻾?⑸땲??
 */

const LOTTEON_GITHUB_CODE_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Code.gs';
const LOTTEON_GITHUB_PATCH_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_24_bootstrap_auto_continue.gs';
const LOTTEON_LOADER_VERSION = 'v1.11';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('LOTTEON ?먮룞??)
    .addItem('GitHub 肄붾뱶 ?곌껐 ?뚯뒪??, 'testLotteonGitHubConnection')
    .addSeparator()
    .addItem('??蹂寃쎌궗??諛섏쁺 ?ㅽ뻾', 'runPendingChangesApproval')
    .addItem('????쒕낫?쒕쭔 鍮좊Ⅸ 媛깆떊', 'refreshDashboardFastOnly')
    .addSeparator()
    .addItem('?꾪꽣蹂??곹뭹???덉쟾 媛깆떊 ?쒖옉/?댁뼱?ㅽ뻾', 'runDailyFilterCountsOnceManual')
    .addItem('?꾪꽣蹂??곹뭹???덉쟾 媛깆떊 ?곹깭 ?뺤씤', 'showDailyFilterCountsStatus')
    .addItem('?꾪꽣蹂??곹뭹???덉쟾 媛깆떊 珥덇린??, 'resetDailyFilterCountsSafeState')
    .addItem('?꾪꽣蹂??곹뭹???덉쟾 媛깆떊 ?쒖옉(留ㅼ씪 06:10)', 'startDailyFilterCountsSchedule')
    .addItem('?꾪꽣蹂??곹뭹???덉쟾 媛깆떊 以묒?', 'stopDailyFilterCountsSchedule')
    .addSeparator()
    .addItem('??荑좏뙜?ъ쟾??濡쒓렇 媛깆떊', 'createRetransmitLogSheet')
    .addItem('???듭떖?붿빟+??쒕낫??媛깆떊', 'refreshCoreSummaryAndDashboardWithRetransmitLogDates')
    .addItem('???댁쁺 ?듭떖 10媛쒕쭔 ?쒖떆/遺덊븘????젣', 'showOperationSheetsOnly')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('?ㅼ젙/愿由?)
        .addItem('GitHub 濡쒕뜑 沅뚰븳 ?뱀씤', 'authorizeLotteonLoader')
        .addItem('ScriptApp ?몃━嫄?沅뚰븳 ?뺤씤', 'authorizeLotteonScriptAppScope')
        .addItem('GitHub 肄붾뱶 罹먯떆 珥덇린??, 'clearLotteonGitHubCodeCache')
        .addItem('?꾪꽣蹂??곹뭹???덉쟾 媛깆떊 ?쒖옉(留ㅼ씪 06:10)', 'startDailyFilterCountsSchedule')
        .addItem('?꾪꽣蹂??곹뭹???덉쟾 媛깆떊 以묒?', 'stopDailyFilterCountsSchedule')
        .addItem('?꾪꽣蹂??곹뭹???덉쟾 媛깆떊 珥덇린??, 'resetDailyFilterCountsSafeState')
        .addSeparator()
        .addItem('蹂寃쎄컧吏 湲곕뒫 ?쒖옉', 'startChangeDetection')
        .addItem('蹂寃쎄컧吏 湲곕뒫 以묒?', 'stopChangeDetection')
        .addItem('蹂寃쎄컧吏 ?뚮옒洹?珥덇린??, 'resetChangeDetectionFlags')
    )
    .addSubMenu(
      ui.createMenu('怨좉툒/蹂듦뎄')
        .addItem('遺媛???먮즺 ?앹꽦/?댁뼱?ㅽ뻾', 'generateVatReportsFullSeparated_v622')
        .addItem('?댁쁺 ?듭떖 10媛쒕쭔 ?쒖떆/遺덊븘????젣', 'showOperationSheetsOnly')
        .addItem('?쒗듃 蹂듦뎄: ?꾩껜 ?쒗듃 ?쒖떆', 'showAllSheets')
        .addItem('API ?몄쬆媛????, 'saveApiCredentials')
        .addItem('API ?곌껐 ?뚯뒪??, 'testApiConnection')
    )
    .addToUi();

  ui.createMenu('LOTTEON ?쒖떇')
    .addItem('?대꼫鍮??덉쟾議곗젙 ?쒖옉/?댁뼱?ㅽ뻾', 'runColumnWidthAutoAdjustStep_v623')
    .addItem('?대꼫鍮??덉쟾議곗젙 ?곹깭 ?뺤씤', 'showColumnWidthAutoAdjustStatus_v623')
    .addItem('?대꼫鍮??덉쟾議곗젙 珥덇린??, 'resetColumnWidthAutoAdjust_v623')
    .addSeparator()
    .addItem('湲덉븸 1,000?⑥쐞 ?쇳몴 ?곸슜', 'applyAmountThousandsFormat_v649')
    .addItem('諛깅텇??% ?곸슜', 'applyPercentOneDecimalFormat_v649')
    .addSeparator()
    .addItem('?쒖떆?쒖떇留?鍮좊Ⅸ ?뺣━', 'applyDisplayStandardsOnlyFast_v623')
    .addToUi();
}

function authorizeLotteonLoader() {
  SpreadsheetApp.getActive();
  UrlFetchApp.fetch(LOTTEON_GITHUB_CODE_URL, { muteHttpExceptions: true, followRedirects: true });
  UrlFetchApp.fetch(LOTTEON_GITHUB_PATCH_URL, { muteHttpExceptions: true, followRedirects: true });
  authorizeLotteonScriptAppScope();
  SpreadsheetApp.getUi().alert('GitHub 濡쒕뜑 沅뚰븳 ?뱀씤 ?꾨즺\n\nScriptApp ?몃━嫄?沅뚰븳源뚯? ?뺤씤?덉뒿?덈떎.\n?ㅽ봽?덈뱶?쒗듃瑜??덈줈怨좎묠????硫붾돱瑜??ㅼ떆 ?ㅽ뻾?섏꽭??');
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
    SpreadsheetApp.getUi().alert('ScriptApp ?몃━嫄?沅뚰븳 ?뺤씤 ?꾨즺\n\n?꾩옱 ?꾨줈?앺듃 ?몃━嫄??? ' + triggers.length);
  } catch (e) {}
}

function lotteonDummyTriggerScopeHandler_() {}

function testLotteonGitHubConnection() {
  const bundle = loadLotteonRemoteBundle_();
  const version = detectRemoteVersion_(bundle);
  SpreadsheetApp.getUi().alert(
    'GitHub 肄붾뱶 ?곌껐 ?깃났\n\n' +
    'Code.gs Raw URL:\n' + LOTTEON_GITHUB_CODE_URL + '\n\n' +
    'Patch Raw URL:\n' + LOTTEON_GITHUB_PATCH_URL + '\n\n' +
    '濡쒕뜑 踰꾩쟾: ' + LOTTEON_LOADER_VERSION + '\n' +
    '濡쒕뱶 ?ш린: ' + bundle.length.toLocaleString('ko-KR') + '??n' +
    '踰꾩쟾 異붿젙: ' + version + '\n' +
    '罹먯떆 諛⑹떇: ?ъ슜 ????n' +
    'ScriptApp 沅뚰븳: 濡쒖뺄 loader???ы븿\n' +
    '?꾪꽣 ?먮룞?덉빟 硫붾돱: 蹂듦뎄??n\n' +
    '?꾪꽣蹂??곹뭹??留ㅼ씪 ?먮룞 媛깆떊 ?쒖옉/以묒? 硫붾돱瑜??ъ슜?????덉뒿?덈떎.'
  );
}

function clearLotteonGitHubCodeCache() {
  ['LOTTEON_REMOTE_CODE_BUNDLE_V14','LOTTEON_REMOTE_CODE_BUNDLE_V13','LOTTEON_REMOTE_CODE_BUNDLE'].forEach(function(k){ try { CacheService.getScriptCache().remove(k); } catch(e) {} });
  PropertiesService.getScriptProperties().deleteProperty('LOTTEON_REMOTE_LAST_VERSION');
  SpreadsheetApp.getUi().alert('GitHub 肄붾뱶 罹먯떆瑜?珥덇린?뷀뻽?듬땲??\n\nv1.10? ??⑸웾 肄붾뱶瑜?罹먯떆????ν븯吏 ?딆뒿?덈떎.');
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
  if (code < 200 || code >= 300) throw new Error(label + ' 濡쒕뱶 ?ㅽ뙣 HTTP ' + code + '\n' + url + '\n' + text.slice(0, 500));
  return text;
}

function runRemoteFunctionByName_(functionName) {
  const bundle = loadLotteonRemoteBundle_();
  const safeName = String(functionName || '').trim();
  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(safeName)) throw new Error('?섎せ???⑥닔紐? ' + safeName);
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
  return all && all.length ? all[all.length - 1] : '?뺤씤遺덇?';
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
function resetDailyFilterCountsSafeState() { return runRemoteFunctionByName_('resetDailyFilterCountsSafeState'); }
function generateVatReportsFullSeparated_v622() { return runRemoteFunctionByName_('generateVatReportsFullSeparated_v622'); }
function runColumnWidthAutoAdjustStep_v623() { return runRemoteFunctionByName_('runColumnWidthAutoAdjustStep_v623'); }
function showColumnWidthAutoAdjustStatus_v623() { return runRemoteFunctionByName_('showColumnWidthAutoAdjustStatus_v623'); }
function resetColumnWidthAutoAdjust_v623() { return runRemoteFunctionByName_('resetColumnWidthAutoAdjust_v623'); }
function applyAmountThousandsFormat_v649() { return runRemoteFunctionByName_('applyAmountThousandsFormat_v649'); }
function applyPercentOneDecimalFormat_v649() { return runRemoteFunctionByName_('applyPercentOneDecimalFormat_v649'); }
function applyDisplayStandardsOnlyFast_v623() { return runRemoteFunctionByName_('applyDisplayStandardsOnlyFast_v623'); }

