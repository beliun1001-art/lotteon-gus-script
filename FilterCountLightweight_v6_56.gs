/**
 * v6.56 dedicated filter-count bundle.
 * This file is intentionally self-contained: no Code.gs or bootstrap patch is
 * loaded by filter menu or continuation handlers.
 */
var FILTER_COUNT_LW_VERSION = 'v6.56';
var FILTER_COUNT_LW_STATE_KEY = 'LOTTEON_FILTER_COUNT_SAFE_JOB_STATE';
var FILTER_COUNT_LW_STATUS_SHEET = '필터별_상품수_자동상태';
var FILTER_COUNT_LW_WORK_SHEET = '필터별_상품수_안전작업';
var FILTER_COUNT_LW_FILTER_SHEET = '필터별_상품수';
var FILTER_COUNT_LW_BATCH_SIZE = 10;
var FILTER_COUNT_LW_GUARD_MS = 22000;
var FILTER_COUNT_LW_HEADERS = ['검색필터명','브랜드명','계정번호','쿠팡계정ID','API_totalCount','API_totalPage','이번조회_행수','필터코드','API_최근수집일자','API_필터생성일','API_최근수집일자_필드','API_필터생성일_필드','메모'];
var FILTER_COUNT_LW_ACCOUNT_MAP = { '01': ['227-27-04928','beliun1021'], '02': ['176-71-00758','beliun1021-1'], '03': ['835-58-00765','beliun1023'], '04': ['606-45-93763','beliun1024'] };

function runDailyFilterCountsOnceManual() { return filterCountLightweightStartOrResume_('MANUAL_NOW', false, true); }
function runDailyFilterCountsStart() { return filterCountLightweightStartOrResume_('DAILY_START', true, false); }
function runDailyFilterCountsContinue() { return filterCountLightweightStep_({ source: 'CONTINUATION', showAlert: false }); }
function showDailyFilterCountsStatus() { var state = filterCountLightweightState_() || { version: FILTER_COUNT_LW_VERSION, status: 'IDLE' }; filterCountLightweightWriteStatus_(state); return filterCountLightweightAlertStatus_(state); }
function resetDailyFilterCountsSafeState() { filterCountLightweightDeleteTriggers_(['runDailyFilterCountsContinue']); PropertiesService.getScriptProperties().deleteProperty(FILTER_COUNT_LW_STATE_KEY); filterCountLightweightResetWork_(); var state = { version: FILTER_COUNT_LW_VERSION, status: 'RESET', nextExecutionScheduled: 'N', lastUpdatedAt: filterCountLightweightNow_() }; filterCountLightweightWriteStatus_(state); return state; }

function startDailyFilterCountsSchedule() {
  var props = PropertiesService.getScriptProperties();
  try {
    filterCountLightweightDeleteTriggers_(['runDailyFilterCountsStart','runDailyFilterCountsContinue']);
    ScriptApp.newTrigger('runDailyFilterCountsStart').timeBased().atHour(6).nearMinute(10).everyDays(1).create();
    props.setProperty('LOTTEON_FILTER_COUNT_SAFE_DAILY_ENABLED', 'Y');
    var state = filterCountLightweightState_() || { version: FILTER_COUNT_LW_VERSION };
    state.status = 'DAILY_SCHEDULED'; state.nextExecutionScheduled = 'N'; state.lastUpdatedAt = filterCountLightweightNow_();
    filterCountLightweightSave_(state); filterCountLightweightWriteStatus_(state); return state;
  } catch (e) { return filterCountLightweightTriggerError_(filterCountLightweightState_() || {}, e); }
}
function stopDailyFilterCountsSchedule() { filterCountLightweightDeleteTriggers_(['runDailyFilterCountsStart','runDailyFilterCountsContinue']); var state = filterCountLightweightState_() || { version: FILTER_COUNT_LW_VERSION }; state.status = 'STOPPED'; state.nextExecutionScheduled = 'N'; state.lastUpdatedAt = filterCountLightweightNow_(); filterCountLightweightSave_(state); filterCountLightweightWriteStatus_(state); return state; }

function filterCountLightweightStartOrResume_(source, forceNew, showAlert) {
  var state = filterCountLightweightState_();
  if (forceNew || !filterCountLightweightActive_(state)) {
    filterCountLightweightDeleteTriggers_(['runDailyFilterCountsContinue']); filterCountLightweightResetWork_();
    state = { version: FILTER_COUNT_LW_VERSION, source: source, phase: 'DISCOVER', status: 'STARTING', page: 1, totalFilters: 0, currentIndex: 0, processedCount: 0, successCount: 0, errorCount: 0, lastFilter: '', lastError: '', nextExecutionScheduled: 'N', startedAt: filterCountLightweightNow_(), lastUpdatedAt: filterCountLightweightNow_() };
    filterCountLightweightSave_(state); filterCountLightweightWriteStatus_(state);
  }
  var result = filterCountLightweightStep_({ source: source, showAlert: showAlert });
  return result;
}
function filterCountLightweightStep_(options) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return { done: false, locked: true };
  var started = Date.now();
  try {
    var state = filterCountLightweightState_();
    if (!filterCountLightweightActive_(state)) return { done: true, reason: 'NO_ACTIVE_JOB' };
    state.status = state.phase === 'DISCOVER' ? 'DISCOVER_RUNNING' : 'COUNT_RUNNING'; state.nextExecutionScheduled = 'N'; state.lastUpdatedAt = filterCountLightweightNow_();
    filterCountLightweightSave_(state); filterCountLightweightWriteStatus_(state);
    if (state.phase === 'DISCOVER') filterCountLightweightDiscover_(state, started); else filterCountLightweightCount_(state, started);
    if (state.phase === 'DONE') { state.status = 'DONE'; state.nextExecutionScheduled = 'N'; filterCountLightweightDeleteTriggers_(['runDailyFilterCountsContinue']); }
    else { var scheduled = filterCountLightweightSchedule_(); if (scheduled.ok) { state.status = state.phase === 'DISCOVER' ? 'DISCOVER_WAIT' : 'COUNT_WAIT'; state.nextExecutionScheduled = 'Y'; } else { state.status = 'TRIGGER_PERMISSION_ERROR'; state.nextExecutionScheduled = 'N'; state.lastError = scheduled.error; } }
    state.lastUpdatedAt = filterCountLightweightNow_(); filterCountLightweightSave_(state); filterCountLightweightWriteStatus_(state);
    return { done: state.phase === 'DONE', state: state };
  } catch (e) { var failed = filterCountLightweightState_() || {}; failed.status = 'ERROR'; failed.lastError = String(e && e.message ? e.message : e); failed.nextExecutionScheduled = 'N'; failed.lastUpdatedAt = filterCountLightweightNow_(); filterCountLightweightSave_(failed); filterCountLightweightWriteStatus_(failed); throw e; }
  finally { try { lock.releaseLock(); } catch (ignore) {} }
}
function filterCountLightweightDiscover_(state, started) {
  if (Date.now() - started >= FILTER_COUNT_LW_GUARD_MS) return;
  var page = filterCountLightweightApi_('filterList', { page: String(state.page || 1), searchQuery: { searchKeyword: '롯백', siteId: '', filterGroup: 'all', sort: 'nameAsc' } });
  var data = page.data || page || {}; var items = filterCountLightweightItems_(page); var rows = items.map(filterCountLightweightRow_).filter(Boolean);
  filterCountLightweightAppendWork_(rows); state.lastFilter = rows.length ? rows[rows.length - 1][0] : 'filterList page ' + state.page; state.page = Number(state.page || 1) + 1;
  if (rows.length && (!data.totalPage || state.page <= Number(data.totalPage))) return;
  var all = filterCountLightweightDedupeWork_(); filterCountLightweightReplace_(filterCountLightweightSheet_(FILTER_COUNT_LW_FILTER_SHEET), all); state.phase = all.length ? 'COUNT' : 'DONE'; state.totalFilters = all.length; state.currentIndex = 0; state.processedCount = 0; state.successCount = 0; state.errorCount = 0;
}
function filterCountLightweightCount_(state, started) {
  var sheet = filterCountLightweightSheet_(FILTER_COUNT_LW_FILTER_SHEET), remaining = sheet.getLastRow() - 1 - Number(state.currentIndex || 0);
  if (remaining <= 0) { state.phase = 'DONE'; return; }
  var size = Math.min(FILTER_COUNT_LW_BATCH_SIZE, remaining), start = Number(state.currentIndex || 0) + 2, rows = sheet.getRange(start, 1, size, FILTER_COUNT_LW_HEADERS.length).getValues();
  for (var i = 0; i < rows.length && Date.now() - started < FILTER_COUNT_LW_GUARD_MS; i++) { var name = String(rows[i][0] || '').trim(); if (!name) { state.currentIndex++; state.processedCount++; continue; } state.lastFilter = name; state.currentIndex++; state.processedCount++; try { var data = filterCountLightweightApi_('productList', { page: '1', searchQuery: { siteId: '', searchType: 'filterName', searchKeyword: name, condition: '', sort: 'dateDesc' } }).data || {}; rows[i][4] = Number(data.totalCount || 0); rows[i][5] = Number(data.totalPage || 0); rows[i][6] = 0; rows[i][12] = 'v6.56 productList API_totalCount'; state.successCount++; } catch (e) { state.errorCount++; state.lastError = name + ': ' + String(e && e.message ? e.message : e); rows[i][12] = String(rows[i][12] || '') + ' / v6.56 오류: ' + String(e && e.message ? e.message : e).slice(0, 250); } }
  sheet.getRange(start, 1, rows.length, FILTER_COUNT_LW_HEADERS.length).setValues(rows); if (state.currentIndex >= state.totalFilters) state.phase = 'DONE';
}
function filterCountLightweightApi_(endpoint, payload) { var props = PropertiesService.getScriptProperties(), base = String(props.getProperty('THE_MANGO_BASE_URL') || 'https://tmg2007.cafe24.com').replace(/\/+$/, ''), key = String(props.getProperty('THE_MANGO_API_KEY') || '').trim(), sender = String(props.getProperty('THE_MANGO_SENDER') || '').trim(); if (!key || !sender) throw new Error('API 인증값이 없습니다.'); var response = UrlFetchApp.fetch(base + '/api/' + endpoint, { method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + key, 'X-API-SENDER': sender }, payload: JSON.stringify(payload), muteHttpExceptions: true }); if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) throw new Error(endpoint + ' HTTP_' + response.getResponseCode()); return JSON.parse(response.getContentText()); }
function filterCountLightweightItems_(response) { var d = response && response.data || response || {}; return (d.filters || d.items || d.data && (d.data.filters || d.data.items) || []).filter(function(v) { return v && typeof v === 'object'; }); }
function filterCountLightweightRow_(item) { var name = String(item.filterName || item.name || item.searchFilterName || '').trim(), m = name.match(/^롯백_(\d{2})_?(.*)$/); if (!m || !m[2] || !FILTER_COUNT_LW_ACCOUNT_MAP[m[1]]) return null; var account = FILTER_COUNT_LW_ACCOUNT_MAP[m[1]], count = Number(item.itemCount || item.totalCount || item.productCount || 0); return [name, m[2], account[0], account[1], count, '', count, m[1], item.updateDate || '', item.createDate || '', '', '', 'v6.56 filterList']; }
function filterCountLightweightState_() { var raw = PropertiesService.getScriptProperties().getProperty(FILTER_COUNT_LW_STATE_KEY); try { return raw ? JSON.parse(raw) : null; } catch (e) { return null; } }
function filterCountLightweightSave_(state) { PropertiesService.getScriptProperties().setProperty(FILTER_COUNT_LW_STATE_KEY, JSON.stringify(state || {})); }
function filterCountLightweightActive_(state) { return !!(state && (state.phase === 'DISCOVER' || state.phase === 'COUNT')); }
function filterCountLightweightSheet_(name) { var ss = SpreadsheetApp.getActive(), sheet = ss.getSheetByName(name); return sheet || ss.insertSheet(name); }
function filterCountLightweightResetWork_() { var sheet = filterCountLightweightSheet_(FILTER_COUNT_LW_WORK_SHEET); sheet.clearContents(); sheet.getRange(1,1,1,FILTER_COUNT_LW_HEADERS.length).setValues([FILTER_COUNT_LW_HEADERS]); }
function filterCountLightweightAppendWork_(rows) { if (!rows.length) return; var sheet = filterCountLightweightSheet_(FILTER_COUNT_LW_WORK_SHEET); sheet.getRange(Math.max(2,sheet.getLastRow()+1),1,rows.length,FILTER_COUNT_LW_HEADERS.length).setValues(rows); }
function filterCountLightweightDedupeWork_() { var sheet = filterCountLightweightSheet_(FILTER_COUNT_LW_WORK_SHEET); if (sheet.getLastRow() < 2) return []; var map = {}; sheet.getRange(2,1,sheet.getLastRow()-1,FILTER_COUNT_LW_HEADERS.length).getValues().forEach(function(row){ if (row[0]) map[String(row[0])] = row; }); return Object.keys(map).sort().map(function(k){ return map[k]; }); }
function filterCountLightweightReplace_(sheet, rows) { sheet.clearContents(); sheet.getRange(1,1,1,FILTER_COUNT_LW_HEADERS.length).setValues([FILTER_COUNT_LW_HEADERS]); if (rows.length) sheet.getRange(2,1,rows.length,FILTER_COUNT_LW_HEADERS.length).setValues(rows); sheet.setFrozenRows(1); }
function filterCountLightweightWriteStatus_(s) { var sheet = filterCountLightweightSheet_(FILTER_COUNT_LW_STATUS_SHEET), rows = [['항목','값'],['버전',s.version || FILTER_COUNT_LW_VERSION],['상태',s.status || 'IDLE'],['전체 필터 수',s.totalFilters || 0],['현재 index',s.currentIndex || 0],['처리 완료 수',s.processedCount || 0],['성공 수',s.successCount || 0],['오류 수',s.errorCount || 0],['다음 실행 예약 여부',s.nextExecutionScheduled || 'N'],['마지막 처리 필터',s.lastFilter || ''],['마지막 오류',s.lastError || ''],['최종 갱신',s.lastUpdatedAt || '']]; sheet.clearContents(); sheet.getRange(1,1,rows.length,2).setValues(rows); sheet.setFrozenRows(1); }
function filterCountLightweightSchedule_() { try { filterCountLightweightDeleteTriggers_(['runDailyFilterCountsContinue']); ScriptApp.newTrigger('runDailyFilterCountsContinue').timeBased().after(60000).create(); return { ok: true }; } catch (e) { return { ok: false, error: '트리거 권한 오류: ' + String(e && e.message ? e.message : e) }; } }
function filterCountLightweightDeleteTriggers_(names) { (ScriptApp.getProjectTriggers() || []).forEach(function(t){ try { if (names.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); } catch (e) {} }); }
function filterCountLightweightTriggerError_(state, error) { state.status = 'TRIGGER_PERMISSION_ERROR'; state.lastError = '트리거 권한 오류: ' + String(error && error.message ? error.message : error); state.nextExecutionScheduled = 'N'; state.lastUpdatedAt = filterCountLightweightNow_(); filterCountLightweightSave_(state); filterCountLightweightWriteStatus_(state); return state; }
function filterCountLightweightNow_() { return new Date().toISOString(); }
function filterCountLightweightAlertStatus_(state) { try { SpreadsheetApp.getUi().alert('필터별_상품수 안전 갱신 상태\n\n상태: ' + state.status + '\n진행: ' + (state.currentIndex || 0) + ' / ' + (state.totalFilters || 0)); } catch (e) {} return state; }

