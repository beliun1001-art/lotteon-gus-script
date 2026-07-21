/**
 * v6.52 - Issue #3 safe column-width runner.
 *
 * Self-contained width-only path: no aggregation, API, dashboard/VAT rebuild,
 * whole-sheet formatting, or autoResizeColumn(s).
 */
var LOTTEON_SAFE_WIDTH_VERSION = 'v6.52';
var LOTTEON_SAFE_WIDTH_JOB_KEY = 'LOTTEON_SAFE_WIDTH_JOB_STATE';
var LOTTEON_SAFE_WIDTH_HANDLER = 'runColumnWidthAutoAdjustStep_v623';
var LOTTEON_SAFE_WIDTH_STATUS_SHEET = '열너비_자동조정상태';
var LOTTEON_SAFE_WIDTH_SHEETS_PER_TICK = 2;
var LOTTEON_SAFE_WIDTH_SAMPLE_ROWS = 80;
var LOTTEON_SAFE_WIDTH_MAX_COLUMNS = 30;
var LOTTEON_SAFE_WIDTH_TICK_LIMIT_MS = 20000;
var LOTTEON_SAFE_WIDTH_TRIGGER_DELAY_MS = 60000;

// Keep the loader-facing entrypoints stable while replacing their implementation.
runColumnWidthAutoAdjustStep_v623 = function() {
  return runSafeColumnWidthStep_v652_();
};

showColumnWidthAutoAdjustStatus_v623 = function() {
  var state = getSafeWidthState_v652_();
  var ss = state && state.spreadsheetId ? SpreadsheetApp.openById(state.spreadsheetId) : SpreadsheetApp.getActive();
  if (!state) {
    state = emptySafeWidthState_v652_(ss);
    state.status = 'not_started';
  }
  writeSafeWidthStatus_v652_(ss, state, '상태 확인');
  safeWidthAlert_v652_(
    '열너비 안전조정 상태\n\n' +
    '상태: ' + state.status + '\n' +
    '현재 위치: ' + Number(state.nextIndex || 0) + ' / ' + Number((state.sheetNames || []).length) + '\n' +
    '다음 시트: ' + (nextSafeWidthSheetName_v652_(state) || '없음') + '\n' +
    '성공/건너뜀: ' + Number(state.processedCount || 0) + ' / ' + Number(state.skippedCount || 0) + '\n' +
    '자동 이어실행: ' + (state.nextRunScheduled ? '예약됨' : '대기 없음') + '\n' +
    '최근 결과: ' + (state.lastResult || '없음')
  );
  return state;
};

resetColumnWidthAutoAdjust_v623 = function() {
  var props = PropertiesService.getScriptProperties();
  [
    LOTTEON_SAFE_WIDTH_JOB_KEY,
    'LOTTEON_COLWIDTH_V623_NEXT_INDEX',
    'LOTTEON_COLWIDTH_V623_LAST_RESULT',
    'LOTTEON_COLWIDTH_V624_AUTO_RUNNING',
    'LOTTEON_COLWIDTH_V624_LAST_UPDATED',
    'LOTTEON_COLUMN_WIDTH_BATCH_STATE_V623'
  ].forEach(function(key) { props.deleteProperty(key); });
  clearSafeWidthTriggers_v652_();
  var ss = SpreadsheetApp.getActive();
  var state = emptySafeWidthState_v652_(ss);
  state.status = 'reset';
  writeSafeWidthStatus_v652_(ss, state, '신규/이전 상태 key 초기화');
  safeWidthAlert_v652_('열너비 안전조정 상태와 예약 트리거를 초기화했습니다.');
  return { ok: true, reset: true };
};

function runSafeColumnWidthStep_v652_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return { ok: false, busy: true };
  try {
    var state = getSafeWidthState_v652_();
    if (!state || state.status === 'done' || state.status === 'failed' || state.status === 'reset') {
      state = startSafeWidthJob_v652_();
    }
    return continueSafeWidthJob_v652_(state);
  } finally {
    lock.releaseLock();
  }
}

function startSafeWidthJob_v652_() {
  clearSafeWidthTriggers_v652_();
  var ss = SpreadsheetApp.getActive();
  var targetNames = getSafeWidthTargetSheets_v652_(ss).map(function(sheet) { return sheet.getName(); });
  var state = emptySafeWidthState_v652_(ss);
  state.status = targetNames.length ? 'running' : 'done';
  state.sheetNames = targetNames;
  state.startedAt = new Date().toISOString();
  state.updatedAt = state.startedAt;
  state.lastResult = targetNames.length ? '대상 선정 완료' : '처리할 운영 시트 없음';
  saveSafeWidthState_v652_(state);
  // Persist and expose status before the first sheet is touched.
  writeSafeWidthStatus_v652_(ss, state, '작업 초기화');
  return state;
}

function continueSafeWidthJob_v652_(state) {
  var started = Date.now();
  var ss = SpreadsheetApp.openById(state.spreadsheetId);
  clearSafeWidthTriggers_v652_();
  state.nextRunScheduled = false;
  var attempted = 0;

  while (state.nextIndex < state.sheetNames.length && attempted < LOTTEON_SAFE_WIDTH_SHEETS_PER_TICK) {
    if (attempted > 0 && Date.now() - started >= LOTTEON_SAFE_WIDTH_TICK_LIMIT_MS) break;
    var sheetName = state.sheetNames[state.nextIndex];
    state.currentSheet = sheetName;
    state.status = 'running';
    state.updatedAt = new Date().toISOString();
    state.lastResult = '처리 시작: ' + sheetName;
    saveSafeWidthState_v652_(state);
    writeSafeWidthStatus_v652_(ss, state, '시트 처리 전 상태 저장');

    try {
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.isSheetHidden() || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
        state.skippedCount += 1;
        state.lastResult = '건너뜀(없음/숨김/빈 시트): ' + sheetName;
      } else {
        var result = applySafeWidthsToSheet_v652_(sheet);
        state.processedCount += 1;
        state.lastResult = '완료: ' + sheetName + ' / 열 ' + result.columns + ' / 샘플행 ' + result.sampleRows;
      }
    } catch (e) {
      state.skippedCount += 1;
      state.lastError = sheetName + ': ' + String(e && e.message ? e.message : e);
      state.lastResult = '오류 건너뜀: ' + state.lastError;
    }

    state.nextIndex += 1;
    attempted += 1;
    state.currentSheet = '';
    state.updatedAt = new Date().toISOString();
    saveSafeWidthState_v652_(state);
    writeSafeWidthStatus_v652_(ss, state, '시트 처리 결과 저장');
  }

  if (state.nextIndex >= state.sheetNames.length) {
    state.status = 'done';
    state.nextRunScheduled = false;
    state.updatedAt = new Date().toISOString();
    state.lastResult = '전체 완료: 성공 ' + state.processedCount + ', 건너뜀 ' + state.skippedCount;
    saveSafeWidthState_v652_(state);
    clearSafeWidthTriggers_v652_();
    writeSafeWidthStatus_v652_(ss, state, '작업 완료');
    safeWidthAlert_v652_(state.lastResult);
    return { ok: true, done: true, state: state };
  }

  scheduleSafeWidthTrigger_v652_(state);
  writeSafeWidthStatus_v652_(ss, state, '다음 tick 예약');
  return { ok: true, done: false, state: state };
}

function applySafeWidthsToSheet_v652_(sheet) {
  var lastRow = sheet.getLastRow();
  var columnCount = Math.min(sheet.getLastColumn(), LOTTEON_SAFE_WIDTH_MAX_COLUMNS);
  var sampleRows = Math.min(Math.max(lastRow - 1, 0), LOTTEON_SAFE_WIDTH_SAMPLE_ROWS);
  var values = sheet.getRange(1, 1, sampleRows + 1, columnCount).getDisplayValues();
  var headers = values[0] || [];

  for (var c = 0; c < columnCount; c++) {
    var header = normalizeSafeWidthText_v652_(headers[c]);
    var maxLength = 0;
    for (var r = 1; r < values.length; r++) {
      var length = displayLengthForWidth_v652_(values[r][c]);
      if (length > maxLength) maxLength = length;
    }
    var width = safeWidthForColumn_v652_(header, maxLength);
    sheet.setColumnWidth(c + 1, width);
  }

  // Header-only styling is bounded to one row and does not touch data formats.
  sheet.getRange(1, 1, 1, columnCount).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  return { columns: columnCount, sampleRows: sampleRows };
}

function getSafeWidthTargetSheets_v652_(ss) {
  var preferred = [
    '대시보드', '브랜드별_마진율', '미정산_쿠팡계정별', '쿠팡재전송_로그',
    '부가세_신고자료', '부가세_상품별', '부가세_기간별', '사업자별_계정별_써머리',
    '사업자번호_매핑검증', '필터별_상품수', '매출데이터_붙여넣기',
    '매출데이터_정리', '통합_브랜드요약', '수동입력/API검증', '검색필터_리포트'
  ];
  var seen = {};
  var result = [];
  preferred.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || seen[name]) return;
    seen[name] = true;
    if (sheet.isSheetHidden()) return;
    if (name === LOTTEON_SAFE_WIDTH_STATUS_SHEET) return;
    result.push(sheet);
  });
  return result;
}

function safeWidthForColumn_v652_(header, dataLength) {
  var h = normalizeSafeWidthText_v652_(header).replace(/\s/g, '');
  if (isUrlOrLongTextHeader_v652_(h)) return clampSafeWidth_v652_(Math.max(180, dataLength * 7 + 24), 180, /URL|링크/i.test(h) ? 260 : 300);
  if (isIdentifierHeaderSafeWidth_v652_(h)) return clampSafeWidth_v652_(Math.max(120, dataLength * 8 + 20), 120, 160);
  if (isDateHeaderSafeWidth_v652_(h)) return clampSafeWidth_v652_(Math.max(75, dataLength * 8 + 18), 70, 90);
  if (isPercentHeaderSafeWidth_v652_(h)) return clampSafeWidth_v652_(Math.max(75, dataLength * 8 + 18), 70, 90);
  if (isAmountHeaderSafeWidth_v652_(h)) return clampSafeWidth_v652_(Math.max(90, dataLength * 8 + 18), 80, 120);
  var fallbackLength = dataLength || Math.min(displayLengthForWidth_v652_(header), 12);
  return clampSafeWidth_v652_(fallbackLength * 8 + 24, 60, 180);
}

function isIdentifierHeaderSafeWidth_v652_(h) {
  return /주문번호|상품번호|마켓주문번호|마켓상품번호|계정ID|아이디|ID$|코드|사업자등록번호/i.test(h) && !/금액|매출|정산|매입/.test(h);
}
function isUrlOrLongTextHeader_v652_(h) { return /URL|링크|상품명|메모|비고|사유|주소|옵션/i.test(h); }
function isDateHeaderSafeWidth_v652_(h) { return /날짜|일자|주문일|결제일|등록일|갱신일|date/i.test(h); }
function isPercentHeaderSafeWidth_v652_(h) { return /률|비율|마진|rate|%/i.test(h) && !/금액/.test(h); }
function isAmountHeaderSafeWidth_v652_(h) { return /금액|매출|정산|매입|이익|수수료|부가세|공급가액|비용|price|cost|fee|vat/i.test(h); }

function displayLengthForWidth_v652_(value) {
  var text = normalizeSafeWidthText_v652_(value);
  var length = 0;
  for (var i = 0; i < text.length; i++) length += text.charCodeAt(i) > 255 ? 1.7 : 1;
  return Math.min(length, 40);
}

function normalizeSafeWidthText_v652_(value) { return String(value == null ? '' : value).replace(/[\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim(); }
function clampSafeWidth_v652_(value, min, max) { return Math.max(min, Math.min(max, Math.round(value))); }

function emptySafeWidthState_v652_(ss) {
  return {
    version: LOTTEON_SAFE_WIDTH_VERSION,
    status: 'new',
    spreadsheetId: ss.getId(),
    sheetNames: [],
    nextIndex: 0,
    currentSheet: '',
    processedCount: 0,
    skippedCount: 0,
    nextRunScheduled: false,
    lastResult: '',
    lastError: '',
    startedAt: '',
    updatedAt: new Date().toISOString()
  };
}

function nextSafeWidthSheetName_v652_(state) { return state.sheetNames && state.nextIndex < state.sheetNames.length ? state.sheetNames[state.nextIndex] : ''; }
function getSafeWidthState_v652_() { var raw = PropertiesService.getScriptProperties().getProperty(LOTTEON_SAFE_WIDTH_JOB_KEY); if (!raw) return null; try { return JSON.parse(raw); } catch (e) { return null; } }
function saveSafeWidthState_v652_(state) { PropertiesService.getScriptProperties().setProperty(LOTTEON_SAFE_WIDTH_JOB_KEY, JSON.stringify(state)); }

function scheduleSafeWidthTrigger_v652_(state) {
  clearSafeWidthTriggers_v652_();
  ScriptApp.newTrigger(LOTTEON_SAFE_WIDTH_HANDLER).timeBased().after(LOTTEON_SAFE_WIDTH_TRIGGER_DELAY_MS).create();
  state.nextRunScheduled = true;
  state.updatedAt = new Date().toISOString();
  saveSafeWidthState_v652_(state);
}

function clearSafeWidthTriggers_v652_() {
  try {
    ScriptApp.getProjectTriggers().forEach(function(trigger) {
      try { if (trigger.getHandlerFunction() === LOTTEON_SAFE_WIDTH_HANDLER) ScriptApp.deleteTrigger(trigger); } catch (e) {}
    });
  } catch (e) {}
}

function writeSafeWidthStatus_v652_(ss, state, memo) {
  var sheet = ss.getSheetByName(LOTTEON_SAFE_WIDTH_STATUS_SHEET) || ss.insertSheet(LOTTEON_SAFE_WIDTH_STATUS_SHEET);
  var rows = [
    ['항목','값','메모'],
    ['버전', state.version || LOTTEON_SAFE_WIDTH_VERSION, 'Issue #3 안전 실행기'],
    ['상태', state.status || '', memo || ''],
    ['처리 위치', Number(state.nextIndex || 0) + ' / ' + Number((state.sheetNames || []).length), '다음: ' + (nextSafeWidthSheetName_v652_(state) || '없음')],
    ['현재 시트', state.currentSheet || '', '처리 전에 기록'],
    ['성공 시트', Number(state.processedCount || 0), ''],
    ['건너뜀 시트', Number(state.skippedCount || 0), '실패 포함'],
    ['다음 실행 예정', state.nextRunScheduled ? 'Y' : 'N', ''],
    ['상태 저장 key', LOTTEON_SAFE_WIDTH_JOB_KEY, 'ScriptProperties'],
    ['최근 결과', state.lastResult || '', ''],
    ['마지막 오류', state.lastError || '', ''],
    ['갱신시각', new Date().toISOString(), '']
  ];
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#d9eaf7');
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 240);
  sheet.setColumnWidth(3, 260);
}

function safeWidthAlert_v652_(message) { try { SpreadsheetApp.getUi().alert(message); } catch (e) {} }

