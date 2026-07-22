/** v6.59 Issue #18: keep dashboard fast path synchronous; continue lifecycle once. */
var LOTTEON_V659_LIFECYCLE_STATE_KEY = 'LOTTEON_V659_BRAND_LIFECYCLE_STATE';
var LOTTEON_V659_LIFECYCLE_HANDLER = 'continueBrandLifecycleDashboard_v659_';
var LOTTEON_V659_LIFECYCLE_GUARD_MS = 45000;
var LOTTEON_V659_LIFECYCLE_ROWS_PER_TICK = 200;
var LOTTEON_V659_TIMING_KEY = 'LOTTEON_V659_DASHBOARD_TIMING';
var __v659Timing_ = null;
var __baseRefreshDashboardFastOnly_v659_ = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;
var __baseDashboardCore_v659_ = typeof __baseRebuildDashboardSingleSource_v658_ === 'function' ? __baseRebuildDashboardSingleSource_v658_ : (typeof rebuildDashboardSingleSource_v628_ === 'function' ? rebuildDashboardSingleSource_v628_ : null);
var __baseSalesAgg_v659_ = typeof buildSingleSourceSalesAgg_v628_ === 'function' ? buildSingleSourceSalesAgg_v628_ : null;
var __baseFilterAgg_v659_ = typeof aggregateFiltersByBrand_v611_ === 'function' ? aggregateFiltersByBrand_v611_ : null;

function v659Timed_(name, fn) { var started = Date.now(); var value = fn(); if (__v659Timing_) __v659Timing_.steps[name] = (__v659Timing_.steps[name] || 0) + (Date.now() - started); return value; }
if (__baseSalesAgg_v659_) buildSingleSourceSalesAgg_v628_ = function() { return v659Timed_('singleSourceSalesAggMs', function() { return __baseSalesAgg_v659_.apply(this, arguments); }.bind(this)); };
if (__baseFilterAgg_v659_) aggregateFiltersByBrand_v611_ = function() { return v659Timed_('filterAggMs', function() { return __baseFilterAgg_v659_.apply(this, arguments); }.bind(this)); };
// Bypass v6.58's immediate append wrapper. The original v6.28 core dashboard stays synchronous.
if (__baseDashboardCore_v659_) rebuildDashboardSingleSource_v628_ = function() { return v659Timed_('coreDashboardMs', function() { return __baseDashboardCore_v659_.apply(this, arguments); }.bind(this)); };
if (__baseRefreshDashboardFastOnly_v659_) refreshDashboardFastOnly = function() {
  __v659Timing_ = { startedAt:new Date().toISOString(), startedMs:Date.now(), steps:{} };
  try {
    var result = __baseRefreshDashboardFastOnly_v659_.apply(this, arguments);
    __v659Timing_.steps.fastDashboardTotalMs = Date.now() - __v659Timing_.startedMs;
    scheduleBrandLifecycleContinuation_v659_(SpreadsheetApp.getActive(), __v659Timing_);
    return result;
  } finally { __v659Timing_ = null; }
};

function getLifecycleState_v659_() { var text = PropertiesService.getScriptProperties().getProperty(LOTTEON_V659_LIFECYCLE_STATE_KEY); try { return text ? JSON.parse(text) : null; } catch (e) { return null; } }
function saveLifecycleState_v659_(state) { PropertiesService.getScriptProperties().setProperty(LOTTEON_V659_LIFECYCLE_STATE_KEY, JSON.stringify(state)); return state; }
function clearLifecycleTriggers_v659_() { ScriptApp.getProjectTriggers().forEach(function(trigger) { if (trigger.getHandlerFunction() === LOTTEON_V659_LIFECYCLE_HANDLER) ScriptApp.deleteTrigger(trigger); }); }
function scheduleLifecycleTrigger_v659_() { clearLifecycleTriggers_v659_(); ScriptApp.newTrigger(LOTTEON_V659_LIFECYCLE_HANDLER).timeBased().after(1000).create(); }
function scheduleBrandLifecycleContinuation_v659_(ss, timing) {
  clearLifecycleTriggers_v659_();
  var state = { status:'pending', spreadsheetId:ss.getId(), requestedAt:new Date().toISOString(), timing:timing || {}, lifecycleRows:0, nextIndex:0, startRow:0, lifecycleMs:0, lastError:'' };
  saveLifecycleState_v659_(state); PropertiesService.getScriptProperties().setProperty(LOTTEON_V659_TIMING_KEY, JSON.stringify(timing || {})); scheduleLifecycleTrigger_v659_();
  return state;
}
function continueBrandLifecycleDashboard_v659_() {
  var state = getLifecycleState_v659_(); if (!state || state.status === 'done') { clearLifecycleTriggers_v659_(); return { skipped:true }; }
  clearLifecycleTriggers_v659_(); var started = Date.now();
  try {
    var ss = SpreadsheetApp.openById(state.spreadsheetId), sales = buildSingleSourceSalesAgg_v628_(), filters = aggregateFiltersByBrand_v611_();
    var result = writeBrandLifecycleChunk_v659_(ss, sales, filters, state, started);
    state.lifecycleRows = result.rows; state.lifecycleMs += Date.now() - started; state.lastError = '';
    if (result.done) { state.status = 'done'; state.completedAt = new Date().toISOString(); saveLifecycleState_v659_(state); clearLifecycleTriggers_v659_(); }
    else { state.status = 'pending'; saveLifecycleState_v659_(state); scheduleLifecycleTrigger_v659_(); }
    return result;
  } catch (e) {
    state.status = 'failed'; state.lastError = String(e && e.message ? e.message : e); state.failedAt = new Date().toISOString(); saveLifecycleState_v659_(state); clearLifecycleTriggers_v659_(); throw e;
  }
}
function writeBrandLifecycleChunk_v659_(ss, salesAgg, filterAgg, state, started) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD); if (!sheet) return { rows:0, done:true };
  var rules = loadLifecycleRules_v658_(ss), rows = buildBrandLifecycleRows_v658_(salesAgg || {}, filterAgg || {}, lifecycleToday_v658_(), rules);
  var headers = ['구분','브랜드명','대표검색필터명','계정ID','수집수','최초수집일','최초수집일 출처','경과일','매출품목수','순수매출액','30일환산 매출품목수','30일환산 순수매출액','운영구분','운영구분 사유','메모'];
  if (!state.startRow) { state.startRow = sheet.getLastRow() + 2; sheet.getRange(state.startRow,1,1,headers.length).setValues([headers]).setBackground('#d9eaf7').setFontWeight('bold'); }
  if (Date.now() - started >= LOTTEON_V659_LIFECYCLE_GUARD_MS) return { rows:rows.length, done:false, guarded:true };
  var from = Number(state.nextIndex || 0), to = Math.min(rows.length, from + LOTTEON_V659_LIFECYCLE_ROWS_PER_TICK), chunk = rows.slice(from, to).map(function(r) { return ['브랜드운영',r.brand,r.filterName,r.accountId,r.collectCount,r.firstDate,r.firstSource,r.elapsed == null ? '' : r.elapsed,r.products,r.sales,r.products30,r.sales30,r.status,r.reason,r.memo]; });
  if (chunk.length) { sheet.getRange(state.startRow + 1 + from,1,chunk.length,headers.length).setValues(chunk); sheet.getRange(state.startRow + 1 + from,10,chunk.length,1).setNumberFormat('#,##0'); sheet.getRange(state.startRow + 1 + from,11,chunk.length,1).setNumberFormat('0.0'); sheet.getRange(state.startRow + 1 + from,12,chunk.length,1).setNumberFormat('#,##0'); }
  state.nextIndex = to;
  if (to < rows.length) return { rows:rows.length, done:false, nextIndex:to };
  applyLifecycleConditionalFormats_v658_(sheet, state.startRow + 1, rows.length);
  return { rows:rows.length, done:true, expansion:rows.filter(function(r){return r.status==='확장';}).length, maintain:rows.filter(function(r){return r.status==='유지';}).length, exit:rows.filter(function(r){return r.status==='퇴출';}).length };
}

