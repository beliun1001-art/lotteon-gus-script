/**
 * LOTTEON v6.55 safe filter-count runner.
 *
 * The filter count menu deliberately does not call dashboard, VAT, sales or
 * retransmit routines.  It only refreshes ?꾪꽣蹂??곹뭹??API_totalCount.
 */
var LOTTEON_PATCH_V655_SAFE_FILTER_COUNT_RUNNER_LOADED = true;
var LOTTEON_FILTER_COUNT_SAFE_VERSION = 'v6.55';
var LOTTEON_FILTER_COUNT_SAFE_STATE_KEY = 'LOTTEON_FILTER_COUNT_SAFE_JOB_STATE';
var LOTTEON_FILTER_COUNT_SAFE_STATUS_SHEET = '?꾪꽣蹂??곹뭹???먮룞?곹깭';
var LOTTEON_FILTER_COUNT_SAFE_WORK_SHEET = '?꾪꽣蹂??곹뭹???덉쟾?묒뾽';
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
    writeSafeFilterCountStatus_v655_(getSafeFilterCountState_v655_() || { status: 'LOCKED', lastError: '?ㅻⅨ ?덉쟾 媛깆떊 ?묒뾽???ㅽ뻾 以묒엯?덈떎.' });
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
        state.lastError = continuation.error || '?댁뼱?ㅽ뻾 ?몃━嫄곕? ?덉빟?섏? 紐삵뻽?듬땲??';
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
    if (options.showAlert) SpreadsheetApp.getUi().alert('?꾪꽣蹂??곹뭹???덉쟾 媛깆떊 ?ㅻ쪟\n\n' + message);
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
      result.done ? '?꾪꽣蹂??곹뭹???덉쟾 媛깆떊 ?꾨즺' : '?꾪꽣蹂??곹뭹???덉쟾 媛깆떊???쒖옉/?댁뼱?ㅽ뻾?덉뒿?덈떎.',
      '?곹깭: ' + current.status + '\n泥섎━: ' + current.processedCount + ' / ' + current.totalFilters + '\n?ㅻ쪟: ' + current.errorCount,
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
      rows[i][12] = 'v6.55 productList API_totalCount ?덉쟾 媛깆떊';
    }
  }
  sheet.getRange(startIndex + 2, 1, rows.length, CONFIG.HEADERS.FILTERS.length).setValues(rows);
  if (state.currentIndex >= state.totalFilters) state.phase = 'DONE';
}

function showDailyFilterCountsStatus() {
  var state = getSafeFilterCountState_v655_() || { version: LOTTEON_FILTER_COUNT_SAFE_VERSION, status: 'IDLE' };
  writeSafeFilterCountStatus_v655_(state);
  SpreadsheetApp.getUi().alert(
    '?꾪꽣蹂??곹뭹???덉쟾 媛깆떊 ?곹깭\n\n' +
    '?곹깭: ' + (state.status || 'IDLE') + '\n' +
    '吏꾪뻾: ' + (state.currentIndex || 0) + ' / ' + (state.totalFilters || 0) + '\n' +
    '?깃났/?ㅻ쪟: ' + (state.successCount || 0) + ' / ' + (state.errorCount || 0) + '\n' +
    '?ㅼ쓬 ?ㅽ뻾 ?덉빟: ' + (state.nextExecutionScheduled || 'N') + '\n' +
    '留덉?留??꾪꽣: ' + (state.lastFilter || '') + '\n' +
    '留덉?留??ㅻ쪟: ' + (state.lastError || '')
  );
  return state;
}

function resetDailyFilterCountsSafeState() {
  deleteSafeFilterCountContinuationTriggers_v655_();
  PropertiesService.getScriptProperties().deleteProperty(LOTTEON_FILTER_COUNT_SAFE_STATE_KEY);
  resetSafeFilterCountWork_v655_();
  var state = { version: LOTTEON_FILTER_COUNT_SAFE_VERSION, status: 'RESET', lastUpdatedAt: now_(), nextExecutionScheduled: 'N' };
  writeSafeFilterCountStatus_v655_(state);
  SpreadsheetApp.getUi().alert('?꾪꽣蹂??곹뭹???덉쟾 媛깆떊 ?곹깭瑜?珥덇린?뷀뻽?듬땲??');
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
    SpreadsheetApp.getUi().alert('?꾪꽣蹂??곹뭹???덉쟾 媛깆떊??留ㅼ씪 06:10 ?꾪썑濡??덉빟?덉뒿?덈떎.');
  } catch (e) {
    var msg = '?몃━嫄?沅뚰븳 ?ㅻ쪟: ' + String(e && e.message ? e.message : e) + '\nApps Script?먯꽌 ScriptApp ?몃━嫄?沅뚰븳???뱀씤?????ㅼ떆 ?ㅽ뻾?섏꽭??';
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
    SpreadsheetApp.getUi().alert('?꾪꽣蹂??곹뭹???덉쟾 媛깆떊 ?덉빟??以묒??덉뒿?덈떎.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('?몃━嫄?沅뚰븳 ?ㅻ쪟: ' + String(e && e.message ? e.message : e));
  }
}

function scheduleSafeFilterCountContinuation_v655_() {
  try {
    deleteSafeFilterCountContinuationTriggers_v655_();
    ScriptApp.newTrigger('runDailyFilterCountsContinue').timeBased().after(60 * 1000).create();
    return { scheduled: true, error: '' };
  } catch (e) {
    return { scheduled: false, error: '?몃━嫄?沅뚰븳 ?ㅻ쪟: ' + String(e && e.message ? e.message : e) };
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
function appendSafeFilterCountErrorMemo_v655_(memo, error) { return String(memo || '').slice(0, 350) + ' / v6.55 ?ㅻ쪟: ' + String(error || '').slice(0, 300); }
function writeSafeFilterCountStatus_v655_(state) {
  var sheet = ensureSheet_(SpreadsheetApp.getActive(), LOTTEON_FILTER_COUNT_SAFE_STATUS_SHEET);
  var rows = [
    ['??ぉ', '媛?], ['踰꾩쟾', state.version || LOTTEON_FILTER_COUNT_SAFE_VERSION], ['?곹깭', state.status || 'IDLE'], ['?꾩껜 ?꾪꽣 ??, state.totalFilters || 0], ['?꾩옱 index', state.currentIndex || 0], ['泥섎━ ?꾨즺 ??, state.processedCount || 0], ['?깃났 ??, state.successCount || 0], ['?ㅻ쪟 ??, state.errorCount || 0], ['?ㅼ쓬 ?ㅽ뻾 ?덉빟 ?щ?', state.nextExecutionScheduled || 'N'], ['留덉?留?泥섎━ ?꾪꽣', state.lastFilter || ''], ['留덉?留??ㅻ쪟', state.lastError || ''], ['理쒖쥌 媛깆떊', state.lastUpdatedAt || '']
  ];
  sheet.clearContents(); sheet.getRange(1, 1, rows.length, 2).setValues(rows); sheet.setFrozenRows(1);
}

