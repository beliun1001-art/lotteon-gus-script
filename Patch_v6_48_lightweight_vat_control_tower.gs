/**
 * v6.48 - Issue #1 lightweight VAT control tower.
 *
 * This file is deliberately self-contained.  The VAT entrypoint below does not
 * call any of the historical sales aggregation or VAT wrapper functions.
 */
var LOTTEON_V648_VERSION = 'v6.48';
var LOTTEON_V648_JOB_KEY = 'LOTTEON_V648_LIGHT_VAT_JOB_STATE';
var LOTTEON_V648_HANDLER = 'generateVatReportsFullSeparated_v622';
var LOTTEON_V648_SOURCE_SHEET = '매출데이터_붙여넣기';
var LOTTEON_V648_DETAIL_SHEET = '부가세_신고자료';
var LOTTEON_V648_STATUS_SHEET = '부가세_생성상태';
var LOTTEON_V648_CHUNK_SIZE = 500;
var LOTTEON_V648_MAX_COL = 29; // A:AC only
var LOTTEON_V648_DELAY_MS = 60000;

/** Menu/trigger entrypoint. First call initializes state; later calls resume. */
generateVatReportsFullSeparated_v622 = function() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: false, busy: true };
  try {
    var state = getVatState_v648_();
    if (!state || state.status === 'done' || state.status === 'failed') {
      return startVatJob_v648_();
    }
    return continueVatJob_v648_(state);
  } finally {
    lock.releaseLock();
  }
};

function startVatJob_v648_() {
  clearVatTriggers_v648_();
  var ss = SpreadsheetApp.getActive();
  var source = ss.getSheetByName(LOTTEON_V648_SOURCE_SHEET);
  if (!source) throw new Error(LOTTEON_V648_SOURCE_SHEET + ' 시트를 찾을 수 없습니다.');

  deleteExcludedVatSheets_v648_(ss);
  initializeVatDetail_v648_(ss);
  var state = {
    version: LOTTEON_V648_VERSION,
    status: 'running',
    phase: 'detail',
    spreadsheetId: ss.getId(),
    sourceRow: 2,
    sourceLastRow: source.getLastRow(),
    writtenRows: 0,
    skippedRows: 0,
    accountMissingRows: 0,
    nextRunScheduled: false,
    startedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    lastError: ''
  };
  saveVatState_v648_(state);
  scheduleVatTrigger_v648_(state);
  writeVatStatus_v648_(ss, state, '초기화 완료; 첫 배치 예약');
  toastVat_v648_(ss, '부가세 자료 경량 생성이 시작되었습니다. 첫 배치를 예약했습니다.');
  return { ok: true, scheduled: true, state: state };
}

function continueVatJob_v648_(state) {
  var ss = openVatSpreadsheet_v648_(state);
  try {
    state.nextRunScheduled = false;
    if (state.phase === 'detail') return runVatDetailBatch_v648_(ss, state);
    if (state.phase === 'summaries') return finishVatSummaries_v648_(ss, state);
    throw new Error('알 수 없는 배치 단계: ' + state.phase);
  } catch (e) {
    state.status = 'failed';
    state.nextRunScheduled = false;
    state.updatedAt = new Date().toISOString();
    state.lastError = String(e && e.message ? e.message : e);
    saveVatState_v648_(state);
    clearVatTriggers_v648_();
    writeVatStatus_v648_(ss, state, '오류로 중단');
    throw e;
  }
}

function runVatDetailBatch_v648_(ss, state) {
  var source = ss.getSheetByName(LOTTEON_V648_SOURCE_SHEET);
  if (!source) throw new Error(LOTTEON_V648_SOURCE_SHEET + ' 시트를 찾을 수 없습니다.');
  var lastRow = Math.max(Number(state.sourceLastRow || 0), source.getLastRow());
  var startRow = Math.max(2, Number(state.sourceRow || 2));

  if (startRow > lastRow) {
    state.phase = 'summaries';
    state.updatedAt = new Date().toISOString();
    saveVatState_v648_(state);
    scheduleVatTrigger_v648_(state);
    writeVatStatus_v648_(ss, state, '상세 완료; 요약 생성 예약');
    return { ok: true, done: false, state: state };
  }

  var maxCol = Math.min(LOTTEON_V648_MAX_COL, Math.max(7, source.getLastColumn()));
  var count = Math.min(LOTTEON_V648_CHUNK_SIZE, lastRow - startRow + 1);
  var headers = source.getRange(1, 1, 1, maxCol).getValues()[0];
  var values = source.getRange(startRow, 1, count, maxCol).getValues();
  var indexes = vatHeaderIndexes_v648_(headers);
  var output = [];
  var skipped = 0;
  var missing = 0;

  values.forEach(function(row, offset) {
    var result = vatDetailRow_v648_(row, indexes, startRow + offset);
    if (!result.row) {
      skipped += 1;
      if (result.accountMissing) missing += 1;
      return;
    }
    if (result.accountMissing) missing += 1;
    output.push(result.row);
  });

  if (output.length) {
    var detail = ss.getSheetByName(LOTTEON_V648_DETAIL_SHEET);
    detail.getRange(2 + Number(state.writtenRows || 0), 1, output.length, output[0].length).setValues(output);
  }
  state.sourceRow = startRow + count;
  state.writtenRows = Number(state.writtenRows || 0) + output.length;
  state.skippedRows = Number(state.skippedRows || 0) + skipped;
  state.accountMissingRows = Number(state.accountMissingRows || 0) + missing;
  state.updatedAt = new Date().toISOString();

  if (state.sourceRow > lastRow) state.phase = 'summaries';
  saveVatState_v648_(state);
  scheduleVatTrigger_v648_(state);
  writeVatStatus_v648_(ss, state, state.phase === 'summaries' ? '상세 완료; 요약 생성 예약' : '다음 상세 배치 예약');
  return { ok: true, done: false, state: state };
}

function finishVatSummaries_v648_(ss, state) {
  var detail = ss.getSheetByName(LOTTEON_V648_DETAIL_SHEET);
  var rows = detail && detail.getLastRow() > 1
    ? detail.getRange(2, 1, detail.getLastRow() - 1, vatDetailHeaders_v648_().length).getValues()
    : [];
  buildVatProductSummary_v648_(ss, rows);
  buildBrandMarginSummary_v648_(ss, rows);
  buildAccountSummary_v648_(ss, rows);
  buildMappingDiagnostic_v648_(ss, rows);

  state.status = 'done';
  state.phase = 'done';
  state.nextRunScheduled = false;
  state.updatedAt = new Date().toISOString();
  state.lastError = '';
  saveVatState_v648_(state);
  clearVatTriggers_v648_();
  writeVatStatus_v648_(ss, state, '경량 생성 완료');
  toastVat_v648_(ss, '부가세/계정별 자료 경량 생성이 완료되었습니다.');
  return { ok: true, done: true, state: state };
}

function vatDetailRow_v648_(row, ix, sourceRow) {
  // D is the only account source. Preserve the displayed value as-is except whitespace.
  var accountId = cleanVatText_v648_(row[3]);
  var accountMissing = !accountId;
  var status = cleanVatText_v648_(valueAt_v648_(row, ix.status));
  if (/취소|반품|환불/.test(status)) return { row: null, accountMissing: accountMissing };
  var sales = vatNumber_v648_(valueAt_v648_(row, ix.sales));
  if (!sales) return { row: null, accountMissing: accountMissing };

  var settlementActual = vatNumber_v648_(valueAt_v648_(row, ix.settlement));
  var settlement = settlementActual || Math.round(sales * 0.901);
  var purchase = vatNumber_v648_(row[28]); // AC is the purchase source of truth.
  var salesVat = splitVat_v648_(sales);
  var purchaseVat = splitVat_v648_(purchase);
  var fee = sales - settlement;
  var profit = settlement - purchase;
  var payableVat = salesVat.vat - purchaseVat.vat;
  var businessNo = businessNoForMarketId_v648_(accountId);

  return { accountMissing: accountMissing, row: [
    vatDateText_v648_(valueAt_v648_(row, ix.date)),
    accountId,
    businessNo,
    cleanVatText_v648_(valueAt_v648_(row, ix.orderNo)),
    cleanVatText_v648_(valueAt_v648_(row, ix.customer)),
    cleanVatText_v648_(valueAt_v648_(row, ix.brand)),
    cleanVatText_v648_(valueAt_v648_(row, ix.productNo)),
    cleanVatText_v648_(valueAt_v648_(row, ix.productName)),
    vatNumber_v648_(valueAt_v648_(row, ix.quantity)) || 1,
    sales,
    salesVat.supply,
    salesVat.vat,
    settlement,
    fee,
    purchase,
    purchaseVat.supply,
    purchaseVat.vat,
    payableVat,
    profit,
    profit - payableVat,
    accountMissing ? 'D열 마켓아이디 공란 (원본 ' + sourceRow + '행)' : (businessNo ? LOTTEON_V648_VERSION + ' D열 기준' : '사업자번호 미매핑')
  ] };
}

function vatHeaderIndexes_v648_(headers) {
  var normalized = {};
  headers.forEach(function(header, index) { normalized[cleanVatText_v648_(header).replace(/\s/g, '')] = index; });
  function find(names, fallback) {
    for (var i = 0; i < names.length; i++) {
      var key = names[i].replace(/\s/g, '');
      if (Object.prototype.hasOwnProperty.call(normalized, key)) return normalized[key];
    }
    return fallback;
  }
  return {
    date: find(['마켓주문일자','주문일자','결제일자','주문일시'], 0),
    orderNo: find(['마켓주문번호','주문번호','주문ID','주문ID(마켓)'], 2),
    sales: find(['결제금액합계(원)','결제금액합계','결제금액','순수매출액','판매금액'], 6),
    settlement: find(['정산예정금액(원)','정산예정금액','실제정산금액','정산금액'], -1),
    status: find(['주문상태','상태','클레임상태','처리상태'], -1),
    customer: find(['고객명','수령인','수취인','구매자','주문자'], -1),
    brand: find(['브랜드명','브랜드'], -1),
    productNo: find(['마켓상품번호','상품번호','상품코드','판매자상품코드'], 4),
    productName: find(['상품명','상품명(옵션포함)','등록상품명'], -1),
    quantity: find(['판매수량','수량','구매수량'], -1)
  };
}

function vatDetailHeaders_v648_() {
  return ['날짜','쿠팡계정ID','사업자등록번호','주문번호','고객명','브랜드명','상품번호','상품명','판매수량','순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료/비용','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','비고'];
}

function initializeVatDetail_v648_(ss) {
  var sheet = ss.getSheetByName(LOTTEON_V648_DETAIL_SHEET) || ss.insertSheet(LOTTEON_V648_DETAIL_SHEET);
  sheet.clearContents();
  writeTable_v648_(sheet, vatDetailHeaders_v648_(), []);
  sheet.setFrozenRows(1);
}

function deleteExcludedVatSheets_v648_(ss) {
  ['부가세_고객별','부가세_주문번호별','고객주소_구분검증'].forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet && ss.getSheets().length > 1) ss.deleteSheet(sheet);
  });
}

function buildVatProductSummary_v648_(ss, rows) {
  var map = {};
  rows.forEach(function(r) {
    var key = cleanVatText_v648_(r[6]) || cleanVatText_v648_(r[7]) || '상품미확인';
    if (!map[key]) map[key] = [key, r[7], r[5], 0, 0, 0, 0, 0, 0, 0];
    map[key][3] += vatNumber_v648_(r[8]); map[key][4] += vatNumber_v648_(r[9]);
    map[key][5] += vatNumber_v648_(r[10]); map[key][6] += vatNumber_v648_(r[11]);
    map[key][7] += vatNumber_v648_(r[14]); map[key][8] += vatNumber_v648_(r[17]); map[key][9] += vatNumber_v648_(r[19]);
  });
  var out = ss.getSheetByName('부가세_상품별') || ss.insertSheet('부가세_상품별');
  writeTable_v648_(out, ['상품번호','상품명','브랜드명','판매수량','순수매출액','매출공급가액','매출부가세','매입금액','납부예상부가세','부가세반영예상이익'], objectRows_v648_(map));
}

function buildBrandMarginSummary_v648_(ss, rows) {
  var map = {};
  rows.forEach(function(r) {
    var key = cleanVatText_v648_(r[5]) || '브랜드미확인';
    if (!map[key]) map[key] = [key, 0, 0, 0, 0, 0, 0];
    map[key][1] += vatNumber_v648_(r[8]); map[key][2] += vatNumber_v648_(r[9]);
    map[key][3] += vatNumber_v648_(r[12]); map[key][4] += vatNumber_v648_(r[14]);
    map[key][5] += vatNumber_v648_(r[18]); map[key][6] += vatNumber_v648_(r[19]);
  });
  var data = objectRows_v648_(map).map(function(r) { return r.concat([r[2] ? r[5] / r[2] : 0, r[2] ? r[6] / r[2] : 0]); });
  var out = ss.getSheetByName('브랜드별_마진율') || ss.insertSheet('브랜드별_마진율');
  writeTable_v648_(out, ['브랜드명','판매수량','순수매출액','정산기준금액','매입금액','예상이익','부가세반영예상이익','예상이익률','부가세반영이익률'], data);
  if (data.length) out.getRange(2, 8, data.length, 2).setNumberFormat('0.0%');
}

function buildAccountSummary_v648_(ss, rows) {
  var map = {};
  rows.forEach(function(r) {
    var key = cleanVatText_v648_(r[1]) || '계정미확인';
    if (!map[key]) map[key] = { row: [key, r[2], 0, 0, 0, 0, 0, 0, 0, 0], orders: {} };
    if (r[3]) map[key].orders[String(r[3])] = true;
    var x = map[key].row;
    x[3] += vatNumber_v648_(r[8]); x[4] += vatNumber_v648_(r[9]); x[5] += vatNumber_v648_(r[12]);
    x[6] += vatNumber_v648_(r[13]); x[7] += vatNumber_v648_(r[14]); x[8] += vatNumber_v648_(r[18]); x[9] += vatNumber_v648_(r[19]);
  });
  var data = Object.keys(map).sort().map(function(key) { var item = map[key]; item.row[2] = Object.keys(item.orders).length; return item.row; });
  var out = ss.getSheetByName('사업자별_계정별_써머리') || ss.insertSheet('사업자별_계정별_써머리');
  writeTable_v648_(out, ['쿠팡계정ID','사업자등록번호','주문건수','판매수량','순수매출액','정산기준금액','마켓수수료/비용','매입금액','예상이익','부가세반영예상이익'], data);
}

function buildMappingDiagnostic_v648_(ss, rows) {
  var map = {};
  rows.forEach(function(r) {
    var account = cleanVatText_v648_(r[1]);
    var businessNo = cleanVatText_v648_(r[2]);
    var key = account + '|' + businessNo;
    if (!map[key]) map[key] = [account || '계정미확인', businessNo || '사업자번호 미매핑', 0, account ? (businessNo ? '정상' : 'D열 값 미매핑') : 'D열 공란'];
    map[key][2] += 1;
  });
  var out = ss.getSheetByName('사업자번호_매핑검증') || ss.insertSheet('사업자번호_매핑검증');
  writeTable_v648_(out, ['D열 마켓아이디','사업자등록번호','행수','검증결과'], objectRows_v648_(map));
}

function writeTable_v648_(sheet, headers, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground('#d9eaf7').setFontWeight('bold');
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.setFrozenRows(1);
}

function writeVatStatus_v648_(ss, state, memo) {
  var sheet = ss.getSheetByName(LOTTEON_V648_STATUS_SHEET) || ss.insertSheet(LOTTEON_V648_STATUS_SHEET);
  var rows = [
    ['항목','값','메모'],
    ['버전', state.version || LOTTEON_V648_VERSION, 'Issue #1 경량 단일 경로'],
    ['상태', state.status || '', state.phase || ''],
    ['원본 진행행', state.sourceRow || 2, (state.sourceLastRow || 1) + '행까지'],
    ['작성행수', state.writtenRows || 0, LOTTEON_V648_DETAIL_SHEET],
    ['제외행수', state.skippedRows || 0, '취소/반품/환불/매출 0'],
    ['계정미확인 행수', state.accountMissingRows || 0, 'D열 공란만 집계'],
    ['다음 실행 예정 여부', state.nextRunScheduled ? 'Y' : 'N', '1분 후 시간 기반 트리거'],
    ['마지막 오류', state.lastError || '', ''],
    ['상태 저장 key', LOTTEON_V648_JOB_KEY, 'ScriptProperties'],
    ['갱신시각', new Date().toISOString(), memo || '']
  ];
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  sheet.getRange(1, 1, 1, 3).setBackground('#d9eaf7').setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function scheduleVatTrigger_v648_(state) {
  clearVatTriggers_v648_();
  ScriptApp.newTrigger(LOTTEON_V648_HANDLER).timeBased().after(LOTTEON_V648_DELAY_MS).create();
  state.nextRunScheduled = true;
  state.updatedAt = new Date().toISOString();
  saveVatState_v648_(state);
}

function clearVatTriggers_v648_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    try { if (trigger.getHandlerFunction() === LOTTEON_V648_HANDLER) ScriptApp.deleteTrigger(trigger); } catch (e) {}
  });
}

function getVatState_v648_() {
  var raw = PropertiesService.getScriptProperties().getProperty(LOTTEON_V648_JOB_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function saveVatState_v648_(state) {
  PropertiesService.getScriptProperties().setProperty(LOTTEON_V648_JOB_KEY, JSON.stringify(state));
}

function openVatSpreadsheet_v648_(state) {
  return state.spreadsheetId ? SpreadsheetApp.openById(state.spreadsheetId) : SpreadsheetApp.getActive();
}

function businessNoForMarketId_v648_(marketId) {
  var normalized = cleanVatText_v648_(marketId).toLowerCase();
  if (normalized === 'beliun1021' || normalized === '1021') return '227-27-04928';
  if (normalized === 'beliun1023' || normalized === '1023') return '835-58-00765';
  if (normalized === 'beliun1024' || normalized === '1024') return '606-45-93763';
  return '';
}

function objectRows_v648_(map) { return Object.keys(map).sort().map(function(key) { return map[key]; }); }
function valueAt_v648_(row, index) { return index >= 0 && index < row.length ? row[index] : ''; }
function cleanVatText_v648_(value) { return String(value == null ? '' : value).trim(); }
function vatNumber_v648_(value) { if (typeof value === 'number') return value; var n = Number(String(value == null ? '' : value).replace(/[원,%\s]/g, '')); return isNaN(n) ? 0 : n; }
function splitVat_v648_(amount) { var total = Math.round(vatNumber_v648_(amount)); var supply = Math.round(total / 1.1); return { supply: supply, vat: total - supply }; }
function vatDateText_v648_(value) { if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Seoul', 'MM/dd'); return cleanVatText_v648_(value); }
function toastVat_v648_(ss, message) { try { ss.toast(message, 'LOTTEON', 5); } catch (e) {} }
