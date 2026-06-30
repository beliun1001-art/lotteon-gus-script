/**
 * LOTTEON v6.44 VAT detail chunked batch execution
 *
 * 문제:
 * - v6.43에서 상태 시트는 먼저 생성되지만, 1단계 '부가세_신고자료' 자체가 커서 시간초과됩니다.
 * - 따라서 단계 단위 분할만으로 부족하고, 가장 큰 부가세_신고자료 시트를 행 단위 배치로 나눠야 합니다.
 *
 * 수정:
 * - generateVatReportsFullSeparated_v622를 v6.44 배치 실행기로 재정의합니다.
 * - 첫 실행은 상태 시트 생성 + 다음 트리거 예약만 수행합니다.
 * - 부가세_신고자료 단계는 300행씩 끊어서 작성합니다.
 * - 각 배치 시작 전에 부가세_생성상태를 먼저 갱신합니다.
 * - 부가세_신고자료 완료 후 나머지 시트는 단계별로 1개씩 자동 이어실행합니다.
 */

var LOTTEON_PATCH_V644_VAT_DETAIL_CHUNKED_BATCH_LOADED = true;
var LOTTEON_V644_JOB_KEY = 'LOTTEON_V644_VAT_JOB_STATE';
var LOTTEON_V644_HANDLER = 'generateVatReportsFullSeparated_v622';
var LOTTEON_V644_DELAY_MS = 60 * 1000;
var LOTTEON_V644_DETAIL_CHUNK_SIZE = 300;

generateVatReportsFullSeparated_v622 = function() {
  return runVatBatchJob_v644_(true);
};

function runVatBatchJob_v644_(allowUi) {
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  var state = getVatState_v644_();

  if (!state || state.status === 'done' || state.status === 'failed') {
    state = newVatState_v644_();
    saveVatState_v644_(state);
    writeVatStatus_v644_(state, '예약됨', 0, 'v6.44 첫 실행: 상태 시트 생성 후 1번째 배치 예약');
    clearVatTriggers_v644_();
    scheduleVatTrigger_v644_();
    if (allowUi && ui) {
      safeAlert_v644_(ui,
        '부가세 신고자료 생성 예약 완료\n\n' +
        'v6.44 기준: 부가세_신고자료를 300행씩 나눠 자동 생성합니다.\n' +
        '첫 배치는 약 1분 뒤 실행됩니다.\n\n' +
        '진행 상태는 부가세_생성상태 시트에서 확인하세요.'
      );
    }
    return { ok: true, scheduled: true, state: state };
  }

  return runVatBatchOneTick_v644_(state, allowUi);
}

function newVatState_v644_() {
  return {
    status: 'running',
    phase: 'vat_detail',
    detailOffset: 0,
    phaseIndex: 0,
    startedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    completed: [],
    errors: [],
    version: 'v6.44',
    chunkSize: LOTTEON_V644_DETAIL_CHUNK_SIZE
  };
}

function runVatBatchOneTick_v644_(state, allowUi) {
  var started = Date.now();
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  try {
    state.updatedAt = new Date().toISOString();
    saveVatState_v644_(state);

    if (state.phase === 'vat_detail') {
      writeVatStatus_v644_(state, '실행중', 0, '부가세_신고자료 배치 작성 시작: offset ' + state.detailOffset);
      var detailResult = runVatDetailChunk_v644_(state);
      state.updatedAt = new Date().toISOString();

      if (detailResult.done) {
        state.completed.push({ key: 'vat_detail', label: '부가세_신고자료', completedAt: new Date().toISOString(), rows: detailResult.totalRows });
        state.phase = 'other_steps';
        state.phaseIndex = 0;
        state.detailTotalRows = detailResult.totalRows;
        state.detailOffset = detailResult.totalRows;
      } else {
        state.detailOffset = detailResult.nextOffset;
        state.detailTotalRows = detailResult.totalRows;
      }

      saveVatState_v644_(state);
      clearVatTriggers_v644_();
      scheduleVatTrigger_v644_();
      writeVatStatus_v644_(state, detailResult.done ? '진행중' : '배치진행중', Date.now() - started, detailResult.message);
      if (allowUi && ui) safeAlert_v644_(ui, detailResult.message + '\n\n다음 작업은 약 1분 뒤 자동 이어실행됩니다.');
      return { ok: true, done: false, state: state };
    }

    if (state.phase === 'other_steps') {
      var steps = getOtherVatSteps_v644_();
      var step = steps[state.phaseIndex];
      if (!step) {
        state.status = 'done';
        state.updatedAt = new Date().toISOString();
        saveVatState_v644_(state);
        clearVatTriggers_v644_();
        writeVatStatus_v644_(state, '완료', Date.now() - started, '모든 단계 완료');
        if (allowUi && ui) safeAlert_v644_(ui, buildVatDoneMessage_v644_(state));
        return { ok: true, done: true, state: state };
      }

      writeVatStatus_v644_(state, '실행중', 0, '현재 단계 실행 중: ' + step.label);
      var salesAgg = prepareVatSalesAgg_v644_();
      step.fn(salesAgg);
      state.completed.push({ key: step.key, label: step.label, completedAt: new Date().toISOString() });
      state.phaseIndex++;
      state.updatedAt = new Date().toISOString();

      var finished = state.phaseIndex >= steps.length;
      if (finished) {
        state.status = 'done';
        saveVatState_v644_(state);
        clearVatTriggers_v644_();
        writeVatStatus_v644_(state, '완료', Date.now() - started, '모든 단계 완료');
        if (allowUi && ui) safeAlert_v644_(ui, buildVatDoneMessage_v644_(state));
        return { ok: true, done: true, state: state };
      }

      saveVatState_v644_(state);
      clearVatTriggers_v644_();
      scheduleVatTrigger_v644_();
      writeVatStatus_v644_(state, '진행중', Date.now() - started, '다음 단계 자동 예약됨: ' + steps[state.phaseIndex].label);
      if (allowUi && ui) {
        safeAlert_v644_(ui,
          '부가세 신고자료 생성 진행중\n\n' +
          '방금 완료: ' + step.label + '\n' +
          '다음 단계: ' + steps[state.phaseIndex].label + '\n\n' +
          '약 1분 뒤 자동 이어실행됩니다.'
        );
      }
      return { ok: true, done: false, state: state };
    }

    throw new Error('알 수 없는 v6.44 phase: ' + state.phase);
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    state.status = 'failed';
    state.updatedAt = new Date().toISOString();
    state.errors = state.errors || [];
    state.errors.push({ phase: state.phase, detailOffset: state.detailOffset, phaseIndex: state.phaseIndex, message: msg, at: new Date().toISOString() });
    saveVatState_v644_(state);
    clearVatTriggers_v644_();
    writeVatStatus_v644_(state, '오류', Date.now() - started, msg);
    try { log_('patch_v644_vat_batch_error', msg); } catch (ignore) {}
    if (allowUi && ui) safeAlert_v644_(ui, '부가세 신고자료 생성 중 오류가 발생했습니다.\n\n' + msg + '\n\n부가세_생성상태 시트를 확인하세요.');
    throw e;
  }
}

function runVatDetailChunk_v644_(state) {
  var salesAgg = prepareVatSalesAgg_v644_();
  var details = salesAgg.detailRows || [];
  var total = details.length;
  var offset = Number(state.detailOffset || 0);
  var size = Number(state.chunkSize || LOTTEON_V644_DETAIL_CHUNK_SIZE);
  var end = Math.min(offset + size, total);
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('부가세_신고자료') || ss.insertSheet('부가세_신고자료');
  var headers = getVatDetailHeaders_v644_();

  if (offset === 0) {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    try { sheet.getRange(1, 1, 1, headers.length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); } catch (e) {}
  }

  if (end > offset) {
    var rows = [];
    for (var i = offset; i < end; i++) rows.push(vatDetailRowForChunk_v644_(details[i]));
    sheet.getRange(2 + offset, 1, rows.length, headers.length).setValues(rows);
    applyVatDetailChunkFormat_v644_(sheet, 2 + offset, rows.length, headers.length);
  }

  var done = end >= total;
  if (done) {
    try { sheet.setFrozenRows(1); } catch (e) {}
    try { sheet.autoResizeColumns(1, Math.min(headers.length, 8)); } catch (e) {}
  }

  return {
    done: done,
    nextOffset: end,
    totalRows: total,
    message: done
      ? '부가세_신고자료 완료: ' + end + ' / ' + total + '행'
      : '부가세_신고자료 배치 완료: ' + end + ' / ' + total + '행'
  };
}

function getVatDetailHeaders_v644_() {
  return ['날짜','쿠팡계정ID','사업자등록번호','주문번호','고객명','브랜드명','상품번호','상품명','판매수량','순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료/비용','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','비고'];
}

function vatDetailRowForChunk_v644_(d) {
  if (typeof vatDetailRowWithBusinessNo_v640_ === 'function') return vatDetailRowWithBusinessNo_v640_(d);
  var salesSplit = splitVatForChunk_v644_(d.netSales);
  var purchaseSplit = splitVatForChunk_v644_(d.purchase);
  var payableVat = salesSplit.vat - purchaseSplit.vat;
  var accountId = String(d.accountId || '계정미확인').trim();
  var bizNo = typeof getBusinessNoByAccount_v640_ === 'function' ? getBusinessNoByAccount_v640_(accountId) : '';
  var profit = num_v644_(d.estimatedProfit);
  return [d.dateText || '', accountId, bizNo, d.orderNo || '', d.customer || '', d.brand || '', d.productNo || '', d.productName || '', num_v644_(d.qty), num_v644_(d.netSales), salesSplit.supply, salesSplit.vat, num_v644_(d.settlementBasis), num_v644_(d.marketFee), num_v644_(d.purchase), purchaseSplit.supply, purchaseSplit.vat, payableVat, profit, profit - payableVat, bizNo ? 'v6.44 배치생성' : '사업자번호 매핑 확인 필요'];
}

function applyVatDetailChunkFormat_v644_(sheet, startRow, rowCount, colCount) {
  if (!rowCount) return;
  try {
    sheet.getRange(startRow, 1, rowCount, colCount).setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    [1].forEach(function(c){ if(c<=colCount) sheet.getRange(startRow,c,rowCount,1).setNumberFormat('@').setHorizontalAlignment('center'); });
    [9,10,11,12,13,14,15,16,17,18,19,20].forEach(function(c){ if(c<=colCount) sheet.getRange(startRow,c,rowCount,1).setNumberFormat('#,##0').setHorizontalAlignment('right'); });
  } catch (e) {}
}

function getOtherVatSteps_v644_() {
  return [
    { key: 'vat_product', label: '부가세_상품별', fn: function(agg) { buildVatProductSingleSource_v628_(agg); } },
    { key: 'vat_customer', label: '부가세_고객별', fn: function(agg) { buildVatCustomerSingleSource_v628_(agg); } },
    { key: 'vat_order', label: '부가세_주문번호별', fn: function(agg) { buildVatOrderSingleSource_v628_(agg); } },
    { key: 'brand_margin', label: '브랜드별_마진율', fn: function(agg) { if (typeof buildBrandMarginSingleSource_v628_ === 'function') buildBrandMarginSingleSource_v628_(agg); } },
    { key: 'account_summary', label: '사업자별_계정별_써머리', fn: function(agg) { if (typeof buildAccountSummarySheet_v637_ === 'function') buildAccountSummarySheet_v637_(agg); } },
    { key: 'unsettled', label: '미정산_쿠팡계정별', fn: function(agg) { if (typeof buildUnsettledSettlementByAccountSheet_v629_ === 'function') buildUnsettledSettlementByAccountSheet_v629_(agg); } },
    { key: 'validation_amounts', label: '시트별_금액검증', fn: function(agg) { if (typeof buildFinancialValidationSheets_v628_ === 'function') buildFinancialValidationSheets_v628_(agg); } },
    { key: 'diag_account', label: '계정/사업자번호 검증', fn: function(agg) { if (typeof buildMarketIdAccountDiagnostic_v638_ === 'function') buildMarketIdAccountDiagnostic_v638_(agg); if (typeof buildBusinessNoMappingDiagnostic_v640_ === 'function') buildBusinessNoMappingDiagnostic_v640_(agg); } },
    { key: 'diag_customer', label: '고객주소_구분검증', fn: function(agg) { if (typeof buildCustomerAddressGroupingDiagnostics_v629_ === 'function') buildCustomerAddressGroupingDiagnostics_v629_(agg); } },
    { key: 'final_format', label: '최종 최소 서식', fn: function(agg) { if (typeof applyVatAutoContinueFinalFormat_v642_ === 'function') applyVatAutoContinueFinalFormat_v642_(); } }
  ];
}

function prepareVatSalesAgg_v644_() {
  var salesAgg = buildSingleSourceSalesAgg_v628_();
  if (typeof forceSalesAggAccountFromColumnD_v638_ === 'function') salesAgg = forceSalesAggAccountFromColumnD_v638_(salesAgg);
  if (typeof applyCustomerAddressGroups_v629_ === 'function') salesAgg = applyCustomerAddressGroups_v629_(salesAgg);
  if (typeof recalculateProfitRates_v632_ === 'function') salesAgg = recalculateProfitRates_v632_(salesAgg);
  return salesAgg;
}

function writeVatStatus_v644_(state, statusText, elapsedMs, memo) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('부가세_생성상태') || ss.insertSheet('부가세_생성상태');
  var otherSteps = getOtherVatSteps_v644_();
  var next = state.phase === 'vat_detail' ? '부가세_신고자료 ' + (state.detailOffset || 0) + '행부터' : (state.phaseIndex < otherSteps.length ? otherSteps[state.phaseIndex].label : '없음');
  var rows = [
    ['항목','값','메모'],
    ['상태', statusText, 'v6.44 행 단위 배치'],
    ['현재구간', state.phase || '', 'vat_detail 또는 other_steps'],
    ['부가세_신고자료 진행', (state.detailOffset || 0) + ' / ' + (state.detailTotalRows || '?'), '300행씩 작성'],
    ['나머지단계', (state.phaseIndex || 0) + ' / ' + otherSteps.length, '부가세_신고자료 이후 단계'],
    ['다음작업', next, ''],
    ['시작시각', state.startedAt || '', ''],
    ['갱신시각', new Date().toISOString(), ''],
    ['이번 실행 소요초', Math.round((elapsedMs || 0) / 1000), ''],
    ['완료단계', (state.completed || []).map(function(x){ return x.label; }).join(' → '), ''],
    ['오류', (state.errors || []).map(function(x){ return '[' + x.phase + '] ' + x.message; }).join(' / '), ''],
    ['메모', memo || '', '']
  ];
  try {
    sheet.clearContents();
    sheet.getRange(1,1,rows.length,3).setValues(rows);
    sheet.getRange(1,1,1,3).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.autoResizeColumns(1,3);
  } catch (e) {}
}

function getVatState_v644_() {
  var raw = PropertiesService.getScriptProperties().getProperty(LOTTEON_V644_JOB_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}
function saveVatState_v644_(state) { PropertiesService.getScriptProperties().setProperty(LOTTEON_V644_JOB_KEY, JSON.stringify(state)); }
function scheduleVatTrigger_v644_() { ScriptApp.newTrigger(LOTTEON_V644_HANDLER).timeBased().after(LOTTEON_V644_DELAY_MS).create(); }
function clearVatTriggers_v644_() { ScriptApp.getProjectTriggers().forEach(function(t){ try { if (t.getHandlerFunction && t.getHandlerFunction() === LOTTEON_V644_HANDLER) ScriptApp.deleteTrigger(t); } catch(e) {} }); }
function safeAlert_v644_(ui, msg) { try { ui.alert(msg); } catch (e) {} }
function buildVatDoneMessage_v644_(state) { return '부가세 신고자료 생성 완료\n\n기준: v6.44 행 단위 배치 자동 이어실행\n부가세_신고자료: ' + (state.detailTotalRows || 0) + '행\n\n확인 시트: 부가세_신고자료, 사업자별_계정별_써머리, 사업자번호_매핑검증, 부가세_생성상태'; }
function splitVatForChunk_v644_(amount) { if (typeof splitVatSafe_v640_ === 'function') return splitVatSafe_v640_(amount); if (typeof splitVat_v628_ === 'function') return splitVat_v628_(amount); amount = Math.round(num_v644_(amount)); var supply = Math.round(amount / 1.1); return { total: amount, supply: supply, vat: amount - supply }; }
function num_v644_(v) { if (typeof toNumber_ === 'function') return toNumber_(v); if (typeof v === 'number') return v; var n = Number(String(v == null ? '' : v).replace(/₩|,|%/g, '').trim()); return isNaN(n) ? 0 : n; }
