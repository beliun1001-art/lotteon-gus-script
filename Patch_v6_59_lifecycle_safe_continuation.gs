/** v6.59 Issue #18: diagnostic core phases and snapshot-based lifecycle continuation. */
var LOTTEON_V659_VERSION = 'v6.59';
var LOTTEON_V659_STATE_KEY = 'LOTTEON_V659_BRAND_LIFECYCLE_STATE';
var LOTTEON_V659_HANDLER = 'continueBrandLifecycleDashboard_v659_';
var LOTTEON_V659_STATUS_SHEET = '브랜드운영_자동상태';
var LOTTEON_V659_SNAPSHOT_SHEET = '브랜드운영_스냅샷';
var LOTTEON_V659_ROWS_PER_TICK = 200;
var LOTTEON_V659_MAX_NO_PROGRESS = 3;
var LOTTEON_V659_SNAPSHOT_WATCHDOG_DELAY_MS = 60000;
var __baseCoreDashboard_v659_ = typeof __baseRebuildDashboardSingleSource_v658_ === 'function' ? __baseRebuildDashboardSingleSource_v658_ : (typeof rebuildDashboardSingleSource_v628_ === 'function' ? rebuildDashboardSingleSource_v628_ : null);

// Do not let v6.58 append the full lifecycle table during the core path.
if (__baseCoreDashboard_v659_) rebuildDashboardSingleSource_v628_ = function(salesAgg, filterAgg) { return __baseCoreDashboard_v659_.apply(this, arguments); };

function v659Now_() { return new Date().toISOString(); }
function v659RunId_() { return String(Date.now()) + '-' + Math.random().toString(36).slice(2, 9); }
function v659Props_() { return PropertiesService.getScriptProperties(); }
function getLifecycleState_v659_() { try { var text = v659Props_().getProperty(LOTTEON_V659_STATE_KEY); return text ? JSON.parse(text) : null; } catch (e) { return null; } }
function saveLifecycleState_v659_(state) { v659Props_().setProperty(LOTTEON_V659_STATE_KEY, JSON.stringify(state)); return state; }
function lifecycleStateTemplate_v659_(ss) { return { version:LOTTEON_V659_VERSION, runId:v659RunId_(), spreadsheetId:ss.getId(), status:'running', phase:'core_init', steps:{}, lifecycleRows:0, nextIndex:0, snapshotReady:false, snapshotAttempts:0, triggerScheduled:false, noProgressAttempts:0, lastProgressAt:v659Now_(), lastError:'', completedAt:'' }; }
function clearLifecycleTriggers_v659_() { ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === LOTTEON_V659_HANDLER) ScriptApp.deleteTrigger(t); }); }
function scheduleLifecycleTrigger_v659_(state, delayMs) { clearLifecycleTriggers_v659_(); ScriptApp.newTrigger(LOTTEON_V659_HANDLER).timeBased().after(delayMs || 1000).create(); state.triggerScheduled = true; saveLifecycleState_v659_(state); return state; }
function writeLifecycleStatus_v659_(ss, state) { var sheet = ss.getSheetByName(LOTTEON_V659_STATUS_SHEET) || ss.insertSheet(LOTTEON_V659_STATUS_SHEET); var rows = [['항목','값'],['version',state.version],['status',state.status],['phase',state.phase],['runId',state.runId],['singleSourceSalesAggMs',state.steps.singleSourceSalesAggMs || 0],['filterAggMs',state.steps.filterAggMs || 0],['coreDashboardMs',state.steps.coreDashboardMs || 0],['brandMarginMs',state.steps.brandMarginMs || 0],['unsettledMs',state.steps.unsettledMs || 0],['validationMs',state.steps.validationMs || 0],['formatMs',state.steps.formatMs || 0],['lifecycleRows',state.lifecycleRows || 0],['nextIndex',state.nextIndex || 0],['triggerScheduled',state.triggerScheduled ? 'Y' : 'N'],['lastProgressAt',state.lastProgressAt || ''],['lastError',state.lastError || ''],['completedAt',state.completedAt || '']]; sheet.clearContents(); sheet.getRange(1,1,rows.length,2).setValues(rows); sheet.setFrozenRows(1); return state; }
function saveStatus_v659_(ss, state) { saveLifecycleState_v659_(state); writeLifecycleStatus_v659_(ss, state); return state; }
function runPhase_v659_(ss, state, phase, fn) { state.phase = phase + ':start'; state.lastProgressAt = v659Now_(); saveStatus_v659_(ss, state); var started = Date.now(); var result = fn(); state.steps[phase + 'Ms'] = Date.now() - started; state.phase = phase + ':done'; state.lastProgressAt = v659Now_(); saveStatus_v659_(ss, state); return result; }
function useSpreadsheetContext_v659_(ss, fn) { if (SpreadsheetApp.setActiveSpreadsheet) SpreadsheetApp.setActiveSpreadsheet(ss); return fn(); }

function refreshDashboardFastOnly() {
  var lock = LockService.getScriptLock(); if (!lock.tryLock(5000)) return { ok:false, busy:true };
  var ss = SpreadsheetApp.getActive(), state = lifecycleStateTemplate_v659_(ss);
  try {
    clearLifecycleTriggers_v659_(); saveStatus_v659_(ss, state);
    var sales = runPhase_v659_(ss, state, 'singleSourceSalesAgg', function() { return useSpreadsheetContext_v659_(ss, function() { return buildSingleSourceSalesAgg_v628_(); }); });
    var filters = runPhase_v659_(ss, state, 'filterAgg', function() { return useSpreadsheetContext_v659_(ss, function() { return aggregateFiltersByBrand_v611_(); }); });
    runPhase_v659_(ss, state, 'coreDashboard', function() { return rebuildDashboardSingleSource_v628_(sales, filters); });
    runPhase_v659_(ss, state, 'brandMargin', function() { return buildBrandMarginSingleSource_v628_(sales); });
    if (typeof buildUnsettledSettlementByAccountSheet_v616_ === 'function') runPhase_v659_(ss, state, 'unsettled', function() { return buildUnsettledSettlementByAccountSheet_v616_(sales); });
    runPhase_v659_(ss, state, 'validation', function() { return buildFinancialValidationSheets_v628_(sales); });
    if (typeof applyFastOutputSheetFormatting_v620_ === 'function') runPhase_v659_(ss, state, 'format', function() { return applyFastOutputSheetFormatting_v620_(); });
    state.status = 'pending'; state.phase = 'lifecycle_snapshot:queued'; state.lastProgressAt = v659Now_(); scheduleLifecycleTrigger_v659_(state); writeLifecycleStatus_v659_(ss, state);
    try { ss.toast('핵심 대시보드 완료. 브랜드운영 자동 이어실행을 예약했습니다.', 'LOTTEON', 5); } catch (ignore) {}
    return { ok:true, coreDone:true, runId:state.runId };
  } catch (e) { state.status = 'failed'; state.lastError = String(e && e.message ? e.message : e); state.phase = state.phase || 'core:error'; state.triggerScheduled = false; clearLifecycleTriggers_v659_(); saveStatus_v659_(ss, state); throw e;
  } finally { lock.releaseLock(); }
}

function snapshotHeaders_v659_() { return ['구분','브랜드명','대표검색필터명','계정ID','수집수','최초수집일','최초수집일 출처','경과일','매출품목수','순수매출액','30일환산 매출품목수','30일환산 순수매출액','운영구분','운영구분 사유','메모']; }
function snapshotRows_v659_(sales, filters, ss) { var rules = loadLifecycleRules_v658_(ss); return buildBrandLifecycleRows_v658_(sales || {}, filters || {}, lifecycleToday_v658_(), rules).map(function(r) { return ['브랜드운영',r.brand,r.filterName,r.accountId,r.collectCount,r.firstDate,r.firstSource,r.elapsed == null ? '' : r.elapsed,r.products,r.sales,r.products30,r.sales30,r.status,r.reason,r.memo]; }); }
function prepareLifecycleSnapshot_v659_(ss, state) { var sales = useSpreadsheetContext_v659_(ss, function() { return buildSingleSourceSalesAgg_v628_(); }), filters = useSpreadsheetContext_v659_(ss, function() { return aggregateFiltersByBrand_v611_(); }), rows = snapshotRows_v659_(sales, filters, ss), sheet = ss.getSheetByName(LOTTEON_V659_SNAPSHOT_SHEET) || ss.insertSheet(LOTTEON_V659_SNAPSHOT_SHEET), headers = snapshotHeaders_v659_(); sheet.clearContents(); sheet.getRange(1,1,1,headers.length).setValues([headers]); if (rows.length) sheet.getRange(2,1,rows.length,headers.length).setValues(rows); try { sheet.hideSheet(); } catch (e) {} state.lifecycleRows = rows.length; state.snapshotReady = true; state.nextIndex = 0; return rows.length; }
function writeLifecycleChunk_v659_(ss, state) { var source = ss.getSheetByName(LOTTEON_V659_SNAPSHOT_SHEET), dashboard = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD), headers = snapshotHeaders_v659_(); if (!source || !dashboard) throw new Error('lifecycle snapshot/dashboard missing'); if (!state.dashboardStartRow) { state.dashboardStartRow = dashboard.getLastRow() + 2; dashboard.getRange(state.dashboardStartRow,1,1,headers.length).setValues([headers]).setBackground('#d9eaf7').setFontWeight('bold'); } var from = Number(state.nextIndex || 0), count = Math.min(LOTTEON_V659_ROWS_PER_TICK, Math.max(0, state.lifecycleRows - from)); if (count) { var rows = source.getRange(2 + from,1,count,headers.length).getValues(); dashboard.getRange(state.dashboardStartRow + 1 + from,1,count,headers.length).setValues(rows); dashboard.getRange(state.dashboardStartRow + 1 + from,10,count,1).setNumberFormat('#,##0'); dashboard.getRange(state.dashboardStartRow + 1 + from,11,count,1).setNumberFormat('0.0'); dashboard.getRange(state.dashboardStartRow + 1 + from,12,count,1).setNumberFormat('#,##0'); state.nextIndex += count; state.lastProgressAt = v659Now_(); state.noProgressAttempts = 0; } else state.noProgressAttempts += 1; return state.nextIndex >= state.lifecycleRows; }
function currentRunMatches_v659_(token) { var current = getLifecycleState_v659_(); return !!current && current.runId === token; }
function continueBrandLifecycleDashboard_v659_() {
  var lock = LockService.getScriptLock(); if (!lock.tryLock(5000)) { var busyState = getLifecycleState_v659_(); if (busyState && busyState.status !== 'done' && busyState.status !== 'failed') scheduleLifecycleTrigger_v659_(busyState, LOTTEON_V659_SNAPSHOT_WATCHDOG_DELAY_MS); return { ok:false, busy:true, rescheduled:!!busyState }; }
  var state = getLifecycleState_v659_();
  try {
    if (!state || state.status === 'done' || state.status === 'failed') return { skipped:true };
    var token = state.runId, ss = SpreadsheetApp.openById(state.spreadsheetId); clearLifecycleTriggers_v659_(); state.triggerScheduled = false;
    if (!currentRunMatches_v659_(token)) return { skipped:true };
    if (!state.snapshotReady) {
      state.snapshotAttempts = Number(state.snapshotAttempts || 0) + 1;
      if (state.snapshotAttempts > 3) { state.status = 'failed'; state.lastError = 'lifecycle snapshot watchdog 재시도 제한 초과'; saveStatus_v659_(ss, state); return { failed:true }; }
      state.phase = 'lifecycle_snapshot:watchdog'; state.triggerScheduled = false; saveStatus_v659_(ss, state); scheduleLifecycleTrigger_v659_(state, LOTTEON_V659_SNAPSHOT_WATCHDOG_DELAY_MS);
      if (!currentRunMatches_v659_(token)) return { skipped:true };
      runPhase_v659_(ss, state, 'lifecycle_snapshot', function() { return prepareLifecycleSnapshot_v659_(ss, state); });
      if (!currentRunMatches_v659_(token)) return { skipped:true };
      clearLifecycleTriggers_v659_(); state.triggerScheduled = false; saveStatus_v659_(ss, state);
    }
    state.phase = 'lifecycle_publish:start'; saveStatus_v659_(ss, state); var done = writeLifecycleChunk_v659_(ss, state); state.phase = done ? 'lifecycle_publish:done' : 'lifecycle_publish:pending';
    if (!currentRunMatches_v659_(token)) return { skipped:true };
    if (state.noProgressAttempts >= LOTTEON_V659_MAX_NO_PROGRESS) { state.status = 'failed'; state.lastError = 'lifecycle 무진행 재예약 제한 초과'; clearLifecycleTriggers_v659_(); state.triggerScheduled = false; }
    else if (done) { state.status = 'done'; state.completedAt = v659Now_(); applyLifecycleConditionalFormats_v658_(ss.getSheetByName(CONFIG.SHEETS.DASHBOARD), state.dashboardStartRow + 1, state.lifecycleRows); clearLifecycleTriggers_v659_(); state.triggerScheduled = false; }
    else { state.status = 'pending'; scheduleLifecycleTrigger_v659_(state); }
    saveStatus_v659_(ss, state); return { ok:true, done:done, rows:state.lifecycleRows, nextIndex:state.nextIndex, runId:state.runId };
  } catch (e) { if (state) { state.status='failed'; state.lastError=String(e && e.message ? e.message : e); state.triggerScheduled=false; clearLifecycleTriggers_v659_(); try { saveStatus_v659_(SpreadsheetApp.openById(state.spreadsheetId), state); } catch (ignore) {} } throw e;
  } finally { lock.releaseLock(); }
}
