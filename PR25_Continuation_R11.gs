/**
 * PR #25 v6.66 order-key continuation runner R11.
 *
 * Safe design:
 * - Payment evidence is order-level, not source-row-level.
 * - Primary key: 쿠팡계정ID + 마켓주문번호.
 * - Fallback: 마켓주문번호 only, but only when the source order belongs to one account.
 * - A detail order is written only when the source payment value is unique or blank.
 * - Missing / multi-account / multi-payment orders stop before any detail write.
 * - Existing VAT amounts are never changed.
 * - No modal alerts; progress is written to PR25_실행상태.
 *
 * Requires SmokeRunner_v1_15_pr25_tracking_payment.gs (R5 helpers)
 * in the same Apps Script project.
 */
const PR25_R11_VERSION = 'v1.22-PR25-ORDERKEY-CONTINUATION-R11';
const PR25_R11_STATE_KEY = 'PR25_TRACKING_CONTINUATION_R11_STATE';
const PR25_R11_HANDLER = 'continuePr25TrackingContinuationR11';
const PR25_R11_STATUS_SHEET = 'PR25_실행상태';
const PR25_R11_DIAG_SHEET = 'PR25_주문키진단';
const PR25_R11_DELAY_MS = 60 * 1000;

function startPr25TrackingContinuationR11() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: false, busy: true };
  try {
    pr25r11_assertHelpers_();
    const ss = SpreadsheetApp.getActive();
    pr25r11_assertSheets_(ss);
    pr25r11_clearTriggers_();

    const source = ss.getSheetByName('매출데이터_붙여넣기');
    const columns = pr25r11_discoverColumns_(source);
    if (columns.orderNo < 0 || columns.account < 0 || (columns.tracking < 0 && columns.fallback < 0)) {
      throw new Error('원본 주문번호/계정/트래킹 결제수단 열을 찾지 못했습니다.');
    }

    const state = {
      version: PR25_R11_VERSION,
      status: 'running',
      phase: 'backfill',
      spreadsheetId: ss.getId(),
      nextRunScheduled: false,
      startedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      sourceLastColumn: source.getLastColumn(),
      trackingHeader: columns.tracking >= 0 ? columns.headers[columns.tracking] : '',
      trackingIndex: columns.tracking,
      fallbackHeader: columns.fallback >= 0 ? columns.headers[columns.fallback] : '',
      fallbackIndex: columns.fallback,
      detailRows: 0,
      detailOrders: 0,
      sourceMatchedOrders: 0,
      accountFallbackOrders: 0,
      missingOrders: 0,
      multiAccountOrders: 0,
      ambiguousPaymentOrders: 0,
      paymentRows: 0,
      blankPaymentRows: 0,
      orderRows: 0,
      matched: 0,
      nonCard: 0,
      ambiguous: 0,
      noMatch: 0,
      lastError: ''
    };

    pr25r11_saveState_(state);
    pr25r11_schedule_(state);
    pr25r11_writeStatus_(ss, state, '주문번호 기준 결제수단 사전검증 예약');
    ss.toast('PR25 R11 자동 이어실행을 시작했습니다.', 'PR25 R11', 8);
    return { ok: true, scheduled: true, state: state };
  } finally {
    lock.releaseLock();
  }
}

function continuePr25TrackingContinuationR11() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: false, busy: true };
  try {
    const state = pr25r11_getState_();
    if (!state || state.status !== 'running') {
      pr25r11_clearTriggers_();
      return { ok: false, reason: 'NO_RUNNING_STATE' };
    }

    pr25r11_assertHelpers_();
    const ss = SpreadsheetApp.openById(state.spreadsheetId);
    state.nextRunScheduled = false;

    try {
      if (state.phase === 'backfill') {
        const result = pr25r11_backfillByOrder_(ss, state);
        Object.keys(result).forEach(function(key) { state[key] = result[key]; });
        state.phase = 'match';
        state.updatedAt = new Date().toISOString();
        pr25r11_saveState_(state);
        pr25r11_schedule_(state);
        pr25r11_writeStatus_(ss, state, '주문번호 기준 결제수단 백필 완료; 카드 재매칭 예약');
        return { ok: true, done: false, phase: state.phase };
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
        pr25r11_saveState_(state);
        pr25r11_clearTriggers_();
        pr25r11_writeStatus_(ss, state, 'PR25 주문번호 기준 카드 재매칭 완료');
        ss.toast('PR25 R11 카드 재매칭이 완료되었습니다.', 'PR25 R11', 8);
        return { ok: true, done: true };
      }

      throw new Error('알 수 없는 PR25 단계: ' + state.phase);
    } catch (e) {
      state.status = 'failed';
      state.nextRunScheduled = false;
      state.updatedAt = new Date().toISOString();
      state.lastError = String(e && e.message ? e.message : e);
      pr25r11_saveState_(state);
      pr25r11_clearTriggers_();
      pr25r11_writeStatus_(ss, state, '오류로 중단');
      try { ss.toast('PR25 R11 작업이 중단되었습니다. PR25_실행상태를 확인하세요.', 'PR25 R11', 10); } catch (ignore) {}
      throw e;
    }
  } finally {
    lock.releaseLock();
  }
}

function pr25r11_backfillByOrder_(ss, state) {
  const source = ss.getSheetByName('매출데이터_붙여넣기');
  const detail = ss.getSheetByName('부가세_신고자료');
  if (!source || source.getLastRow() < 2) throw new Error('매출데이터_붙여넣기 데이터가 없습니다.');
  if (!detail || detail.getLastRow() < 2) throw new Error('부가세_신고자료 데이터가 없습니다.');

  const sourceValues = source.getDataRange().getDisplayValues();
  const sourceHeaders = sourceValues[0] || [];
  const sourceRows = sourceValues.slice(1);
  const sx = pr25r11_discoverColumns_(source);
  if (sx.orderNo < 0 || sx.account < 0 || (sx.tracking < 0 && sx.fallback < 0)) {
    throw new Error('원본 주문번호/계정/트래킹 결제수단 열을 찾지 못했습니다.');
  }

  state.sourceLastColumn = sourceHeaders.length;
  state.trackingHeader = sx.tracking >= 0 ? sourceHeaders[sx.tracking] : '';
  state.trackingIndex = sx.tracking;
  state.fallbackHeader = sx.fallback >= 0 ? sourceHeaders[sx.fallback] : '';
  state.fallbackIndex = sx.fallback;

  const exactGroups = {};
  const orderGroups = {};

  sourceRows.forEach(function(row, offset) {
    const orderNo = pr25r11_text_(row[sx.orderNo]);
    if (!orderNo) return;
    const account = pr25r11_text_(row[sx.account]);
    const tracking = sx.tracking >= 0 ? pr25r11_text_(row[sx.tracking]) : '';
    const fallback = sx.fallback >= 0 ? pr25r11_text_(row[sx.fallback]) : '';
    const payment = tracking || fallback;

    const exactKey = pr25r11_key_(account, orderNo);
    if (!exactGroups[exactKey]) exactGroups[exactKey] = pr25r11_newGroup_();
    pr25r11_addToGroup_(exactGroups[exactKey], account, payment, offset + 2);

    if (!orderGroups[orderNo]) orderGroups[orderNo] = pr25r11_newGroup_();
    pr25r11_addToGroup_(orderGroups[orderNo], account, payment, offset + 2);
  });

  const detailValues = detail.getDataRange().getValues();
  const detailHeaders = detailValues[0] || [];
  const detailRows = detailValues.slice(1);
  const dx = {
    account: pr25r11_find_(detailHeaders, ['쿠팡계정ID','마켓아이디']),
    orderNo: pr25r11_find_(detailHeaders, ['주문번호','마켓주문번호','주문ID','주문ID(마켓)'])
  };
  if (dx.account < 0 || dx.orderNo < 0) throw new Error('부가세_신고자료 계정ID 또는 주문번호 열을 찾지 못했습니다.');

  let paymentIx = pr25r11_find_(detailHeaders, ['롯데결제수단']);
  if (paymentIx < 0) {
    paymentIx = detailHeaders.length;
    detail.getRange(1, paymentIx + 1).setValue('롯데결제수단');
    detailHeaders.push('롯데결제수단');
  }

  const detailOrderInfo = {};
  detailRows.forEach(function(row, index) {
    const account = pr25r11_text_(row[dx.account]);
    const orderNo = pr25r11_text_(row[dx.orderNo]);
    const key = pr25r11_key_(account, orderNo);
    if (!detailOrderInfo[key]) detailOrderInfo[key] = { account: account, orderNo: orderNo, rowIndexes: [] };
    detailOrderInfo[key].rowIndexes.push(index);
  });

  const resolved = {};
  const diagnostics = [];
  let sourceMatchedOrders = 0;
  let accountFallbackOrders = 0;
  let missingOrders = 0;
  let multiAccountOrders = 0;
  let ambiguousPaymentOrders = 0;

  Object.keys(detailOrderInfo).forEach(function(key) {
    const info = detailOrderInfo[key];
    if (!info.orderNo) {
      missingOrders++;
      diagnostics.push(['MISSING_ORDER_NO', info.account, info.orderNo, '', '', '', info.rowIndexes.length]);
      return;
    }

    let group = exactGroups[key] || null;
    let method = 'ACCOUNT+ORDER';
    if (!group) {
      const orderGroup = orderGroups[info.orderNo] || null;
      if (!orderGroup) {
        missingOrders++;
        diagnostics.push(['SOURCE_ORDER_NOT_FOUND', info.account, info.orderNo, '', '', '', info.rowIndexes.length]);
        return;
      }
      const accounts = Object.keys(orderGroup.accounts);
      if (accounts.length !== 1) {
        multiAccountOrders++;
        diagnostics.push(['MULTI_SOURCE_ACCOUNT', info.account, info.orderNo, accounts.join(' | '), Object.keys(orderGroup.payments).join(' | '), orderGroup.rows.join(','), info.rowIndexes.length]);
        return;
      }
      group = orderGroup;
      method = 'ORDER_ONLY_UNIQUE_ACCOUNT';
      accountFallbackOrders++;
    }

    const payments = Object.keys(group.payments).filter(function(value) { return value !== ''; });
    if (payments.length > 1) {
      ambiguousPaymentOrders++;
      diagnostics.push(['MULTI_PAYMENT', info.account, info.orderNo, Object.keys(group.accounts).join(' | '), payments.join(' | '), group.rows.join(','), info.rowIndexes.length]);
      return;
    }

    const payment = payments.length === 1 ? payments[0] : '';
    resolved[key] = payment;
    sourceMatchedOrders++;
    diagnostics.push([method, info.account, info.orderNo, Object.keys(group.accounts).join(' | '), payment, group.rows.join(','), info.rowIndexes.length]);
  });

  state.detailRows = detailRows.length;
  state.detailOrders = Object.keys(detailOrderInfo).length;
  state.sourceMatchedOrders = sourceMatchedOrders;
  state.accountFallbackOrders = accountFallbackOrders;
  state.missingOrders = missingOrders;
  state.multiAccountOrders = multiAccountOrders;
  state.ambiguousPaymentOrders = ambiguousPaymentOrders;
  pr25r11_saveState_(state);
  pr25r11_writeOrderDiagnostic_(ss, diagnostics);

  if (missingOrders || multiAccountOrders || ambiguousPaymentOrders) {
    throw new Error(
      '주문번호 기준 안전검증 실패: 누락 ' + missingOrders +
      '건 / 복수계정 ' + multiAccountOrders +
      '건 / 복수결제수단 ' + ambiguousPaymentOrders +
      '건. 아무 값도 쓰지 않았습니다.'
    );
  }

  const paymentsToWrite = [];
  let paymentRows = 0;
  let blankPaymentRows = 0;
  detailRows.forEach(function(row) {
    const key = pr25r11_key_(pr25r11_text_(row[dx.account]), pr25r11_text_(row[dx.orderNo]));
    const payment = Object.prototype.hasOwnProperty.call(resolved, key) ? resolved[key] : '';
    paymentsToWrite.push([payment]);
    if (payment) paymentRows++;
    else blankPaymentRows++;
  });

  detail.getRange(2, paymentIx + 1, paymentsToWrite.length, 1).setNumberFormat('@').setValues(paymentsToWrite);
  SpreadsheetApp.flush();

  return {
    detailRows: detailRows.length,
    detailOrders: Object.keys(detailOrderInfo).length,
    sourceMatchedOrders: sourceMatchedOrders,
    accountFallbackOrders: accountFallbackOrders,
    missingOrders: 0,
    multiAccountOrders: 0,
    ambiguousPaymentOrders: 0,
    paymentRows: paymentRows,
    blankPaymentRows: blankPaymentRows
  };
}

function pr25r11_newGroup_() {
  return { accounts: {}, payments: {}, rows: [] };
}

function pr25r11_addToGroup_(group, account, payment, rowNo) {
  group.accounts[account] = true;
  group.payments[payment] = true;
  group.rows.push(rowNo);
}

function pr25r11_discoverColumns_(source) {
  const headers = source.getRange(1, 1, 1, source.getLastColumn()).getDisplayValues()[0] || [];
  let tracking = pr25r11_find_(headers, ['현지트래킹번호','트래킹 번호','트래킹번호','tracking number','trackingnumber']);
  if (tracking < 0) {
    for (let i = 0; i < headers.length; i++) {
      const compact = pr25r11_compact_(headers[i]);
      if (compact.indexOf('트래킹번호') >= 0 || compact.indexOf('trackingnumber') >= 0) {
        tracking = i;
        break;
      }
    }
  }
  return {
    headers: headers,
    account: pr25r11_find_(headers, ['마켓아이디','쿠팡계정ID']),
    orderNo: pr25r11_find_(headers, ['마켓주문번호','주문번호','주문ID','주문ID(마켓)']),
    tracking: tracking,
    fallback: pr25r11_find_(headers, ['결제수단','결제정보','결제방법','카드사','구매결제수단'])
  };
}

function pr25r11_writeOrderDiagnostic_(ss, rows) {
  const headers = ['판정','상세 계정ID','주문번호','원본 계정ID','결제수단 후보','원본 행번호','상세 행수'];
  const sheet = ss.getSheetByName(PR25_R11_DIAG_SHEET) || ss.insertSheet(PR25_R11_DIAG_SHEET);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground('#d9eaf7').setFontWeight('bold');
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.setFrozenRows(1);
}

function pr25r11_assertHelpers_() {
  const required = [
    'pr25r5_groupOrders_','pr25r5_loadHistory_','pr25r5_loadMaster_','pr25r5_canonicalize_',
    'pr25r5_allocate_','pr25r5_writeDiagnostic_','pr25r5_writeSummary_','pr25r5_writeMonthlyPayment_'
  ];
  const missing = required.filter(function(name) {
    try { return typeof eval(name) !== 'function'; } catch (e) { return true; }
  });
  if (missing.length) throw new Error('PR25 R5 helper 누락: ' + missing.join(', '));
}

function pr25r11_assertSheets_(ss) {
  const required = ['매출데이터_붙여넣기','부가세_신고자료','카드사용내역_붙여넣기'];
  const missing = required.filter(function(name) { return !ss.getSheetByName(name); });
  if (missing.length) throw new Error('필수 시트 누락: ' + missing.join(', '));
}

function pr25r11_find_(headers, names) {
  for (let n = 0; n < names.length; n++) {
    const target = pr25r11_compact_(names[n]);
    for (let i = 0; i < headers.length; i++) {
      if (pr25r11_compact_(headers[i]) === target) return i;
    }
  }
  return -1;
}

function pr25r11_text_(value) {
  return String(value == null ? '' : value).trim();
}

function pr25r11_compact_(value) {
  return pr25r11_text_(value).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g, '');
}

function pr25r11_key_(account, orderNo) {
  return pr25r11_text_(account) + '\u001f' + pr25r11_text_(orderNo);
}

function pr25r11_getState_() {
  const raw = PropertiesService.getScriptProperties().getProperty(PR25_R11_STATE_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function pr25r11_saveState_(state) {
  PropertiesService.getScriptProperties().setProperty(PR25_R11_STATE_KEY, JSON.stringify(state));
}

function pr25r11_schedule_(state) {
  pr25r11_clearTriggers_();
  ScriptApp.newTrigger(PR25_R11_HANDLER).timeBased().after(PR25_R11_DELAY_MS).create();
  state.nextRunScheduled = true;
  state.updatedAt = new Date().toISOString();
  pr25r11_saveState_(state);
}

function pr25r11_clearTriggers_() {
  const handlers = {
    continuePr25TrackingContinuationR6: true,
    continuePr25TrackingContinuationR7: true,
    continuePr25TrackingContinuationR8: true,
    continuePr25TrackingContinuationR9: true,
    continuePr25TrackingContinuationR11: true
  };
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    try {
      if (handlers[trigger.getHandlerFunction()]) ScriptApp.deleteTrigger(trigger);
    } catch (e) {}
  });
}

function pr25r11_writeStatus_(ss, state, message) {
  const sheet = ss.getSheetByName(PR25_R11_STATUS_SHEET) || ss.insertSheet(PR25_R11_STATUS_SHEET);
  const rows = [
    ['항목','값'],
    ['버전',state.version || PR25_R11_VERSION],
    ['상태',state.status || ''],
    ['단계',state.phase || ''],
    ['메시지',message || ''],
    ['다음 실행 예약',state.nextRunScheduled ? 'Y' : 'N'],
    ['시작시각',state.startedAt || ''],
    ['갱신시각',state.updatedAt || ''],
    ['행 매핑 기준','쿠팡계정ID + 마켓주문번호 (주문 단위)'],
    ['원본 마지막 열',Number(state.sourceLastColumn || 0)],
    ['트래킹 헤더',state.trackingHeader || ''],
    ['트래킹 열번호',Number(state.trackingIndex) >= 0 ? Number(state.trackingIndex) + 1 : ''],
    ['부가세 상세행',Number(state.detailRows || 0)],
    ['부가세 주문건수',Number(state.detailOrders || 0)],
    ['원본 매칭 주문',Number(state.sourceMatchedOrders || 0)],
    ['주문번호 단독 보정',Number(state.accountFallbackOrders || 0)],
    ['원본 주문 누락',Number(state.missingOrders || 0)],
    ['복수 원본 계정',Number(state.multiAccountOrders || 0)],
    ['복수 결제수단 주문',Number(state.ambiguousPaymentOrders || 0)],
    ['결제수단 입력행',Number(state.paymentRows || 0)],
    ['결제수단 공란행',Number(state.blankPaymentRows || 0)],
    ['최종 주문건수',Number(state.orderRows || 0)],
    ['MATCHED',Number(state.matched || 0)],
    ['NON_CARD',Number(state.nonCard || 0)],
    ['AMBIGUOUS',Number(state.ambiguous || 0)],
    ['NO_MATCH',Number(state.noMatch || 0)],
    ['마지막 오류',state.lastError || '']
  ];
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange(1, 1, 1, 2).setBackground('#d9eaf7').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 680);
  SpreadsheetApp.flush();
}
