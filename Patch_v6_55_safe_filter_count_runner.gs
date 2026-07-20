/**
 * LOTTEON v6.55 safe filter-count runner.
 *
 * The filter count menu deliberately does not call dashboard, VAT, sales or
 * retransmit routines.  It only refreshes 필터별_상품수.API_totalCount.
 */
var LOTTEON_PATCH_V655_SAFE_FILTER_COUNT_RUNNER_LOADED = true;
var LOTTEON_FILTER_COUNT_SAFE_VERSION = 'v6.55';
var LOTTEON_FILTER_COUNT_SAFE_STATE_KEY = 'LOTTEON_FILTER_COUNT_SAFE_JOB_STATE';
var LOTTEON_FILTER_COUNT_SAFE_STATUS_SHEET = '필터별_상품수_자동상태';
var LOTTEON_FILTER_COUNT_SAFE_WORK_SHEET = '필터별_상품수_안전작업';
var LOTTEON_FILTER_COUNT_SAFE_BATCH_SIZE = 10;
var LOTTEON_FILTER_COUNT_SAFE_TIMEOUT_MS = 22000;

function runDailyFilterCountsOnceManual() {
  return startOrResumeSafeFilterCountJob_v655_('MANUAL_NOW', false, true);
}

function runDailyFilterCountsStart() {
  return startOrResumeSafeFilterCountJob_v655_('DAILY_START', true, false);
}

function runDailyFilterCountsContinue() {
  return runDailyFilterCountsStep_({ source: 'CONTINUATION', showAlert: false });
}

function runDailyFilterCountsStep_(options) {
  options = options || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    writeSafeFilterCountStatus_v655_(getSafeFilterCountState_v655_() || { status: 'LOCKED', lastError: '다른 안전 갱신 작업이 실행 중입니다.' });
    return { done: false, locked: true };
  }

  var startedAt = Date.now();
  try {
    var state = getSafeFilterCountState_v655_();
    if (!state || !isSafeFilterCountActive_v655_(state)) {
      return { done: true, reason: 'NO_ACTIVE_JOB' };
    }

    // Status is always written before the first remote API request.
    state.status = state.phase === 'DISCOVER' ? 'DISCOVER_RUNNING' : 'COUNT_RUNNING';
    state.lastUpdatedAt = now_();
    state.nextExecutionScheduled = 'N';
    saveSafeFilterCountState_v655_(state);
    writeSafeFilterCountStatus_v655_(state);

    if (state.phase === 'DISCOVER') {
      runSafeFilterDiscoveryTick_v655_(state, startedAt);
    } else {
      runSafeFilterCountTick_v655_(state, startedAt);
    }

    state.lastUpdatedAt = now_();
    if (state.phase === 'DONE') {
      state.status = 'DONE';
      state.nextExecutionScheduled = 'N';
      deleteSafeFilterCountContinuationTriggers_v655_();
    } else {
      var continuation = scheduleSafeFilterCountContinuation_v655_();
      if (continuation.scheduled) {
        state.status = state.phase === 'DISCOVER' ? 'DISCOVER_WAIT' : 'COUNT_WAIT';
        state.nextExecutionScheduled = 'Y';
      } else {
        state.status = 'TRIGGER_PERMISSION_ERROR';
        state.lastError = continuation.error || '이어실행 트리거를 예약하지 못했습니다.';
        state.nextExecutionScheduled = 'N';
      }
    }
    saveSafeFilterCountState_v655_(state);
    writeSafeFilterCountStatus_v655_(state);
    return { done: state.phase === 'DONE', state: state };
  } catch (e) {
    var message = String(e && e.message ? e.message : e);
    var failed = getSafeFilterCountState_v655_() || {};
    failed.status = 'ERROR';
    failed.lastError = message;
    failed.lastUpdatedAt = now_();
    failed.nextExecutionScheduled = 'N';
    saveSafeFilterCountState_v655_(failed);
    writeSafeFilterCountStatus_v655_(failed);
    deleteSafeFilterCountContinuationTriggers_v655_();
    if (options.showAlert) SpreadsheetApp.getUi().alert('필터별_상품수 안전 갱신 오류\n\n' + message);
    throw e;
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function startOrResumeSafeFilterCountJob_v655_(source, forceNew, showAlert) {
  var state = getSafeFilterCountState_v655_();
  if (forceNew || !isSafeFilterCountActive_v655_(state)) {
    deleteSafeFilterCountContinuationTriggers_v655_();
    resetSafeFilterCountWork_v655_();
    state = {
      version: LOTTEON_FILTER_COUNT_SAFE_VERSION,
      source: source,
      phase: 'DISCOVER',
      status: 'STARTING',
      page: 1,
      totalFilters: 0,
      currentIndex: 0,
      processedCount: 0,
      successCount: 0,
      errorCount: 0,
      lastFilter: '',
      lastError: '',
      startedAt: now_(),
      lastUpdatedAt: now_(),
      nextExecutionScheduled: 'N'
    };
    saveSafeFilterCountState_v655_(state);
    // Explicitly make the starting state visible before the first tick.
    writeSafeFilterCountStatus_v655_(state);
  }
  var result = runDailyFilterCountsStep_({ source: source, showAlert: false });
  if (showAlert) {
    var current = result.state || getSafeFilterCountState_v655_() || state;
    SpreadsheetApp.getUi().alert(
      result.done ? '필터별_상품수 안전 갱신 완료' : '필터별_상품수 안전 갱신을 시작/이어실행했습니다.',
      '상태: ' + current.status + '\n처리: ' + current.processedCount + ' / ' + current.totalFilters + '\n오류: ' + current.errorCount,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
  return result;
}

function runSafeFilterDiscoveryTick_v655_(state, startedAt) {
  if (Date.now() - startedAt >= LOTTEON_FILTER_COUNT_SAFE_TIMEOUT_MS) return;
  var result = fetchFilterListPageForLotteon_(Math.max(1, Number(state.page || 1)));
  var rows = buildFilterSummaryRowsFromFilterListItems_(result.items || []);
  appendSafeFilterCountWorkRows_v655_(rows);
  state.lastFilter = rows.length ? String(rows[rows.length - 1][0] || '') : 'filterList page ' + state.page;
  state.page = Number(state.page || 1) + 1;

  var isLastPage = !rows.length || (result.totalPage && state.page > Number(result.totalPage));
  if (!isLastPage) return;

  var workRows = dedupeSafeFilterCountWorkRows_v655_();
  replaceDataFastLimited_(getSheet_(CONFIG.SHEETS.FILTERS), CONFIG.HEADERS.FILTERS, workRows);
  state.phase = 'COUNT';
  state.totalFilters = workRows.length;
  state.currentIndex = 0;
  state.processedCount = 0;
  state.successCount = 0;
  state.errorCount = 0;
  state.lastFilter = workRows.length ? String(workRows[0][0] || '') : '';
  if (!workRows.length) state.phase = 'DONE';
}

function runSafeFilterCountTick_v655_(state, startedAt) {
  var sheet = getSheet_(CONFIG.SHEETS.FILTERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { state.phase = 'DONE'; return; }
  var startIndex = Math.max(0, Number(state.currentIndex || 0));
  var remaining = lastRow - 1 - startIndex;
  if (remaining <= 0) { state.phase = 'DONE'; return; }
  var count = Math.min(LOTTEON_FILTER_COUNT_SAFE_BATCH_SIZE, remaining);
  var rows = sheet.getRange(startIndex + 2, 1, count, CONFIG.HEADERS.FILTERS.length).getValues();
  var filterNames = rows.map(function(row) { return String(row[0] || '').trim(); }).filter(Boolean);
  var results = fetchProductListTotalCountFast_(filterNames);

  for (var i = 0; i < rows.length; i++) {
    if (Date.now() - startedAt >= LOTTEON_FILTER_COUNT_SAFE_TIMEOUT_MS) break;
    var name = String(rows[i][0] || '').trim();
    if (!name) { state.currentIndex++; state.processedCount++; continue; }
    var item = results[name] || { error: 'EMPTY_API_RESULT', totalCount: 0, totalPage: 0 };
    state.lastFilter = name;
    state.currentIndex++;
    state.processedCount++;
    if (item.error) {
      state.errorCount++;
      state.lastError = name + ': ' + item.error;
      rows[i][12] = appendSafeFilterCountErrorMemo_v655_(rows[i][12], item.error);
      // Keep the previous API_totalCount when this filter fails.
    } else {
      state.successCount++;
      rows[i][4] = Number(item.totalCount || 0);
      rows[i][5] = Number(item.totalPage || 0);
      rows[i][6] = 0;
      rows[i][12] = 'v6.55 productList API_totalCount 안전 갱신';
    }
  }
  sheet.getRange(startIndex + 2, 1, rows.length, CONFIG.HEADERS.FILTERS.length).setValues(rows);
  if (state.currentIndex >= state.totalFilters) state.phase = 'DONE';
}

function showDailyFilterCountsStatus() {
  var state = getSafeFilterCountState_v655_() || { version: LOTTEON_FILTER_COUNT_SAFE_VERSION, status: 'IDLE' };
  writeSafeFilterCountStatus_v655_(state);
  SpreadsheetApp.getUi().alert(
    '필터별_상품수 안전 갱신 상태\n\n' +
    '상태: ' + (state.status || 'IDLE') + '\n' +
    '진행: ' + (state.currentIndex || 0) + ' / ' + (state.totalFilters || 0) + '\n' +
    '성공/오류: ' + (state.successCount || 0) + ' / ' + (state.errorCount || 0) + '\n' +
    '다음 실행 예약: ' + (state.nextExecutionScheduled || 'N') + '\n' +
    '마지막 필터: ' + (state.lastFilter || '') + '\n' +
    '마지막 오류: ' + (state.lastError || '')
  );
  return state;
}

function resetDailyFilterCountsSafeState() {
  deleteSafeFilterCountContinuationTriggers_v655_();
  PropertiesService.getScriptProperties().deleteProperty(LOTTEON_FILTER_COUNT_SAFE_STATE_KEY);
  resetSafeFilterCountWork_v655_();
  var state = { version: LOTTEON_FILTER_COUNT_SAFE_VERSION, status: 'RESET', lastUpdatedAt: now_(), nextExecutionScheduled: 'N' };
  writeSafeFilterCountStatus_v655_(state);
  SpreadsheetApp.getUi().alert('필터별_상품수 안전 갱신 상태를 초기화했습니다.');
}

function startDailyFilterCountsSchedule() {
  var props = PropertiesService.getScriptProperties();
  try {
    deleteSafeFilterCountDailyTriggers_v655_();
    ScriptApp.newTrigger('runDailyFilterCountsStart').timeBased().atHour(6).nearMinute(10).everyDays(1).create();
    props.setProperty('LOTTEON_FILTER_COUNT_SAFE_DAILY_ENABLED', 'Y');
    var state = getSafeFilterCountState_v655_() || { version: LOTTEON_FILTER_COUNT_SAFE_VERSION };
    state.status = 'DAILY_SCHEDULED'; state.lastUpdatedAt = now_(); state.nextExecutionScheduled = 'N';
    saveSafeFilterCountState_v655_(state); writeSafeFilterCountStatus_v655_(state);
    SpreadsheetApp.getUi().alert('필터별_상품수 안전 갱신을 매일 06:10 전후로 예약했습니다.');
  } catch (e) {
    var msg = '트리거 권한 오류: ' + String(e && e.message ? e.message : e) + '\nApps Script에서 ScriptApp 트리거 권한을 승인한 뒤 다시 실행하세요.';
    var failed = getSafeFilterCountState_v655_() || { version: LOTTEON_FILTER_COUNT_SAFE_VERSION };
    failed.status = 'TRIGGER_PERMISSION_ERROR'; failed.lastError = msg; failed.lastUpdatedAt = now_();
    saveSafeFilterCountState_v655_(failed); writeSafeFilterCountStatus_v655_(failed);
    SpreadsheetApp.getUi().alert(msg);
  }
}

function stopDailyFilterCountsSchedule() {
  try {
    deleteSafeFilterCountDailyTriggers_v655_();
    var state = getSafeFilterCountState_v655_() || { version: LOTTEON_FILTER_COUNT_SAFE_VERSION };
    state.status = 'STOPPED'; state.nextExecutionScheduled = 'N'; state.lastUpdatedAt = now_();
    saveSafeFilterCountState_v655_(state); writeSafeFilterCountStatus_v655_(state);
    SpreadsheetApp.getUi().alert('필터별_상품수 안전 갱신 예약을 중지했습니다.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('트리거 권한 오류: ' + String(e && e.message ? e.message : e));
  }
}

function scheduleSafeFilterCountContinuation_v655_() {
  try {
    deleteSafeFilterCountContinuationTriggers_v655_();
    ScriptApp.newTrigger('runDailyFilterCountsContinue').timeBased().after(60 * 1000).create();
    return { scheduled: true, error: '' };
  } catch (e) {
    return { scheduled: false, error: '트리거 권한 오류: ' + String(e && e.message ? e.message : e) };
  }
}

function deleteSafeFilterCountDailyTriggers_v655_() { return deleteSafeFilterCountTriggers_v655_(['runDailyFilterCountsStart', 'runDailyFilterCountsContinue']); }
function deleteSafeFilterCountContinuationTriggers_v655_() { return deleteSafeFilterCountTriggers_v655_(['runDailyFilterCountsContinue']); }
function deleteSafeFilterCountTriggers_v655_(names) {
  var deleted = 0;
  (ScriptApp.getProjectTriggers() || []).forEach(function(trigger) {
    try { if (names.indexOf(trigger.getHandlerFunction()) >= 0) { ScriptApp.deleteTrigger(trigger); deleted++; } } catch (ignore) {}
  });
  return deleted;
}

function getSafeFilterCountState_v655_() {
  var raw = PropertiesService.getScriptProperties().getProperty(LOTTEON_FILTER_COUNT_SAFE_STATE_KEY);
  try { return raw ? JSON.parse(raw) : null; } catch (e) { return null; }
}
function saveSafeFilterCountState_v655_(state) { PropertiesService.getScriptProperties().setProperty(LOTTEON_FILTER_COUNT_SAFE_STATE_KEY, JSON.stringify(state || {})); }
function isSafeFilterCountActive_v655_(state) { return !!(state && (state.phase === 'DISCOVER' || state.phase === 'COUNT')); }
function getSafeFilterCountWorkSheet_v655_() {
  var sheet = ensureSheet_(SpreadsheetApp.getActive(), LOTTEON_FILTER_COUNT_SAFE_WORK_SHEET);
  if (sheet.getLastRow() < 1) sheet.getRange(1, 1, 1, CONFIG.HEADERS.FILTERS.length).setValues([CONFIG.HEADERS.FILTERS]);
  return sheet;
}
function resetSafeFilterCountWork_v655_() {
  var sheet = getSafeFilterCountWorkSheet_v655_();
  sheet.clearContents(); sheet.getRange(1, 1, 1, CONFIG.HEADERS.FILTERS.length).setValues([CONFIG.HEADERS.FILTERS]);
}
function appendSafeFilterCountWorkRows_v655_(rows) {
  if (!rows || !rows.length) return;
  var sheet = getSafeFilterCountWorkSheet_v655_();
  sheet.getRange(Math.max(2, sheet.getLastRow() + 1), 1, rows.length, CONFIG.HEADERS.FILTERS.length).setValues(rows);
}
function dedupeSafeFilterCountWorkRows_v655_() {
  var sheet = getSafeFilterCountWorkSheet_v655_();
  if (sheet.getLastRow() < 2) return [];
  var map = {};
  sheet.getRange(2, 1, sheet.getLastRow() - 1, CONFIG.HEADERS.FILTERS.length).getValues().forEach(function(row) { if (String(row[0] || '').trim()) map[String(row[0]).trim()] = row; });
  return Object.keys(map).sort().map(function(key) { return map[key]; });
}
function appendSafeFilterCountErrorMemo_v655_(memo, error) { return String(memo || '').slice(0, 350) + ' / v6.55 오류: ' + String(error || '').slice(0, 300); }
function writeSafeFilterCountStatus_v655_(state) {
  var sheet = ensureSheet_(SpreadsheetApp.getActive(), LOTTEON_FILTER_COUNT_SAFE_STATUS_SHEET);
  var rows = [
    ['항목', '값'], ['버전', state.version || LOTTEON_FILTER_COUNT_SAFE_VERSION], ['상태', state.status || 'IDLE'], ['전체 필터 수', state.totalFilters || 0], ['현재 index', state.currentIndex || 0], ['처리 완료 수', state.processedCount || 0], ['성공 수', state.successCount || 0], ['오류 수', state.errorCount || 0], ['다음 실행 예약 여부', state.nextExecutionScheduled || 'N'], ['마지막 처리 필터', state.lastFilter || ''], ['마지막 오류', state.lastError || ''], ['최종 갱신', state.lastUpdatedAt || '']
  ];
  sheet.clearContents(); sheet.getRange(1, 1, rows.length, 2).setValues(rows); sheet.setFrozenRows(1);
}

