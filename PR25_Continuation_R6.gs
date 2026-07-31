/**
 * PR #25 v6.66 trigger-based continuation runner R6.
 *
 * Same operating pattern as the existing VAT continuation job:
 * - start once
 * - persist state
 * - continue automatically with 1-minute time triggers
 * - report progress in PR25_실행상태
 * - no modal alert calls
 *
 * Requires SmokeRunner_v1_15_pr25_tracking_payment.gs (R5 helpers)
 * in the same Apps Script project.
 */
const PR25_R6_VERSION = 'v1.17-PR25-CONTINUATION-R6';
const PR25_R6_STATE_KEY = 'PR25_TRACKING_CONTINUATION_R6_STATE';
const PR25_R6_HANDLER = 'continuePr25TrackingContinuationR6';
const PR25_R6_STATUS_SHEET = 'PR25_실행상태';
const PR25_R6_DELAY_MS = 60 * 1000;

/** Run this once. It only initializes state and schedules the first phase. */
function startPr25TrackingContinuationR6() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: false, busy: true };
  try {
    pr25r6_assertHelpers_();
    const ss = SpreadsheetApp.getActive();
    pr25r6_assertSheets_(ss);
    pr25r6_clearTriggers_();

    const state = {
      version: PR25_R6_VERSION,
      status: 'running',
      phase: 'backfill',
      spreadsheetId: ss.getId(),
      nextRunScheduled: false,
      startedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      detailRows: 0,
      paymentRows: 0,
      blankPaymentRows: 0,
      orderRows: 0,
      matched: 0,
      nonCard: 0,
      ambiguous: 0,
      noMatch: 0,
      lastError: ''
    };

    pr25r6_saveState_(state);
    pr25r6_schedule_(state);
    pr25r6_writeStatus_(ss, state, '초기화 완료; 트래킹 결제수단 백필 예약');
    ss.toast('PR25 자동 이어실행을 시작했습니다. PR25_실행상태 시트에서 진행률을 확인하세요.', 'PR25 R6', 8);
    return { ok: true, scheduled: true, state: state };
  } finally {
    lock.releaseLock();
  }
}

/** Time-trigger continuation handler. Do not run repeatedly by hand. */
function continuePr25TrackingContinuationR6() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: false, busy: true };
  try {
    const state = pr25r6_getState_();
    if (!state || state.status !== 'running') {
      pr25r6_clearTriggers_();
      return { ok: false, reason: 'NO_RUNNING_STATE' };
    }

    pr25r6_assertHelpers_();
    const ss = SpreadsheetApp.openById(state.spreadsheetId);
    state.nextRunScheduled = false;

    try {
      if (state.phase === 'backfill') {
        const backfill = pr25r5_backfillTrackingPayment_(ss);
        state.detailRows = Number(backfill.detailRows || 0);
        state.paymentRows = Number(backfill.paymentRows || 0);
        state.blankPaymentRows = Number(backfill.blankPaymentRows || 0);
        state.phase = 'match';
        state.updatedAt = new Date().toISOString();
        pr25r6_saveState_(state);
        pr25r6_schedule_(state);
        pr25r6_writeStatus_(ss, state, '트래킹 결제수단 백필 완료; 카드 재매칭 예약');
        return { ok: true, done: false, phase: state.phase, state: state };
      }

      if (state.phase === 'match') {
        const detail = ss.getSheetByName('부가세_신고자료');
        if (!detail || detail.getLastRow() < 2) throw new Error('부가세_신고자료 데이터가 없습니다.');

        const orders = pr25r5_groupOrders_(detail.getDataRange().getValues());
        const history = pr25r5_loadHistory_(ss);
        const master = pr25r5_loadMaster_(ss);
        const canonical = pr25r5_canonicalize_(history, master);
        const stats = pr25r5_allocate_(orders, canonical, master);

        pr25r5_writeDiagnostic_(ss, orders);
        pr25r5_writeSummary_(ss, orders);
        pr25r5_writeMonthlyPayment_(ss, orders);
        SpreadsheetApp.flush();

        state.orderRows = orders.length;
        state.matched = Number(stats.matched || 0);
        state.nonCard = Number(stats.nonCard || 0);
        state.ambiguous = Number(stats.ambiguous || 0);
        state.noMatch = Number(stats.noMatch || 0);
        state.status = 'done';
        state.phase = 'done';
        state.nextRunScheduled = false;
        state.updatedAt = new Date().toISOString();
        state.lastError = '';
        pr25r6_saveState_(state);
        pr25r6_clearTriggers_();
        pr25r6_writeStatus_(ss, state, 'PR25 트래킹 결제수단 카드 재매칭 완료');
        ss.toast('PR25 카드 재매칭이 완료되었습니다.', 'PR25 R6', 8);
        return { ok: true, done: true, state: state };
      }

      throw new Error('알 수 없는 PR25 단계: ' + state.phase);
    } catch (e) {
      state.status = 'failed';
      state.nextRunScheduled = false;
      state.updatedAt = new Date().toISOString();
      state.lastError = String(e && e.message ? e.message : e);
      pr25r6_saveState_(state);
      pr25r6_clearTriggers_();
      pr25r6_writeStatus_(ss, state, '오류로 중단');
      try { ss.toast('PR25 작업이 중단되었습니다. PR25_실행상태 시트를 확인하세요.', 'PR25 R6', 10); } catch (ignore) {}
      throw e;
    }
  } finally {
    lock.releaseLock();
  }
}

/** Writes the latest persisted state to the status sheet. No modal dialog. */
function showPr25TrackingContinuationStatusR6() {
  const state = pr25r6_getState_();
  const ss = state && state.spreadsheetId
    ? SpreadsheetApp.openById(state.spreadsheetId)
    : SpreadsheetApp.getActive();
  pr25r6_writeStatus_(ss, state || {
    version: PR25_R6_VERSION,
    status: 'not_started',
    phase: '-',
    nextRunScheduled: false,
    startedAt: '',
    updatedAt: '',
    lastError: ''
  }, '현재 상태 확인');
  ss.toast('PR25_실행상태 시트를 갱신했습니다.', 'PR25 R6', 5);
  return state;
}

/** Stops only PR25 R6 triggers and clears the R6 state. Output sheets are preserved. */
function resetPr25TrackingContinuationR6() {
  pr25r6_clearTriggers_();
  PropertiesService.getScriptProperties().deleteProperty(PR25_R6_STATE_KEY);
  const ss = SpreadsheetApp.getActive();
  pr25r6_writeStatus_(ss, {
    version: PR25_R6_VERSION,
    status: 'reset',
    phase: '-',
    nextRunScheduled: false,
    startedAt: '',
    updatedAt: new Date().toISOString(),
    detailRows: 0,
    paymentRows: 0,
    blankPaymentRows: 0,
    orderRows: 0,
    matched: 0,
    nonCard: 0,
    ambiguous: 0,
    noMatch: 0,
    lastError: ''
  }, '상태 초기화 완료');
  ss.toast('PR25 R6 상태와 트리거를 초기화했습니다.', 'PR25 R6', 5);
}

function pr25r6_assertHelpers_() {
  const required = [
    'pr25r5_backfillTrackingPayment_',
    'pr25r5_groupOrders_',
    'pr25r5_loadHistory_',
    'pr25r5_loadMaster_',
    'pr25r5_canonicalize_',
    'pr25r5_allocate_',
    'pr25r5_writeDiagnostic_',
    'pr25r5_writeSummary_',
    'pr25r5_writeMonthlyPayment_'
  ];
  const missing = required.filter(function(name) {
    try { return typeof eval(name) !== 'function'; } catch (e) { return true; }
  });
  if (missing.length) {
    throw new Error('PR25 R5 helper 누락: ' + missing.join(', ') + '. SmokeRunner R5 파일을 같은 프로젝트에 유지하세요.');
  }
}

function pr25r6_assertSheets_(ss) {
  const required = ['매출데이터_붙여넣기', '부가세_신고자료', '카드사용내역_붙여넣기'];
  const missing = required.filter(function(name) { return !ss.getSheetByName(name); });
  if (missing.length) throw new Error('필수 시트 누락: ' + missing.join(', '));
}

function pr25r6_getState_() {
  const raw = PropertiesService.getScriptProperties().getProperty(PR25_R6_STATE_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function pr25r6_saveState_(state) {
  PropertiesService.getScriptProperties().setProperty(PR25_R6_STATE_KEY, JSON.stringify(state));
}

function pr25r6_schedule_(state) {
  pr25r6_clearTriggers_();
  ScriptApp.newTrigger(PR25_R6_HANDLER).timeBased().after(PR25_R6_DELAY_MS).create();
  state.nextRunScheduled = true;
  state.updatedAt = new Date().toISOString();
  pr25r6_saveState_(state);
}

function pr25r6_clearTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    try {
      if (trigger.getHandlerFunction() === PR25_R6_HANDLER) ScriptApp.deleteTrigger(trigger);
    } catch (e) {}
  });
}

function pr25r6_writeStatus_(ss, state, message) {
  const sheet = ss.getSheetByName(PR25_R6_STATUS_SHEET) || ss.insertSheet(PR25_R6_STATUS_SHEET);
  const rows = [
    ['항목', '값'],
    ['버전', state.version || PR25_R6_VERSION],
    ['상태', state.status || ''],
    ['단계', state.phase || ''],
    ['메시지', message || ''],
    ['다음 실행 예약', state.nextRunScheduled ? 'Y' : 'N'],
    ['시작시각', state.startedAt || ''],
    ['갱신시각', state.updatedAt || ''],
    ['부가세 상세행', Number(state.detailRows || 0)],
    ['결제수단 입력행', Number(state.paymentRows || 0)],
    ['결제수단 공란행', Number(state.blankPaymentRows || 0)],
    ['주문건수', Number(state.orderRows || 0)],
    ['MATCHED', Number(state.matched || 0)],
    ['NON_CARD', Number(state.nonCard || 0)],
    ['AMBIGUOUS', Number(state.ambiguous || 0)],
    ['NO_MATCH', Number(state.noMatch || 0)],
    ['마지막 오류', state.lastError || '']
  ];
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange(1, 1, 1, 2).setBackground('#d9eaf7').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 520);
  SpreadsheetApp.flush();
}
