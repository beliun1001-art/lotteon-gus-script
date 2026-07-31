/**
 * PR #25 v6.66 trigger-based continuation runner R8.
 *
 * Safe subset mapping:
 * - VAT financial identity stays A:AC.
 * - Tracking/payment headers are searched across all source columns.
 * - Every existing VAT detail row must match exactly one source row.
 * - Extra source rows are allowed and ignored.
 * - No modal alerts; progress is written to PR25_실행상태.
 *
 * Requires SmokeRunner_v1_15_pr25_tracking_payment.gs (R5 helpers)
 * in the same Apps Script project.
 */
const PR25_R8_VERSION = 'v1.19-PR25-CONTINUATION-R8';
const PR25_R8_STATE_KEY = 'PR25_TRACKING_CONTINUATION_R8_STATE';
const PR25_R8_HANDLER = 'continuePr25TrackingContinuationR8';
const PR25_R8_STATUS_SHEET = 'PR25_실행상태';
const PR25_R8_DELAY_MS = 60 * 1000;
const PR25_R8_CORE_COLS = 29;

function startPr25TrackingContinuationR8() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: false, busy: true };
  try {
    pr25r8_assertHelpers_();
    const ss = SpreadsheetApp.getActive();
    pr25r8_assertSheets_(ss);
    pr25r8_clearTriggers_();

    const source = ss.getSheetByName('매출데이터_붙여넣기');
    const found = pr25r8_discoverPaymentColumns_(source);
    if (found.trackingIndex < 0 && found.fallbackIndex < 0) {
      throw new Error(
        '원본 전체 ' + found.lastColumn + '열에서 트래킹 번호 또는 결제수단 열을 찾지 못했습니다. 관련 헤더: ' +
        (found.relatedHeaders.length ? found.relatedHeaders.join(' | ') : '없음')
      );
    }

    const state = {
      version: PR25_R8_VERSION,
      status: 'running',
      phase: 'backfill',
      spreadsheetId: ss.getId(),
      nextRunScheduled: false,
      startedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      sourceLastColumn: found.lastColumn,
      trackingIndex: found.trackingIndex,
      trackingHeader: found.trackingHeader,
      fallbackIndex: found.fallbackIndex,
      fallbackHeader: found.fallbackHeader,
      sourceEligibleRows: 0,
      detailRows: 0,
      matchedDetailRows: 0,
      ignoredSourceRows: 0,
      unmatchedDetailRows: 0,
      paymentRows: 0,
      blankPaymentRows: 0,
      orderRows: 0,
      matched: 0,
      nonCard: 0,
      ambiguous: 0,
      noMatch: 0,
      lastError: ''
    };

    pr25r8_saveState_(state);
    pr25r8_schedule_(state);
    pr25r8_writeStatus_(ss, state, '상세행 100% 매핑 방식으로 트래킹 결제수단 백필 예약');
    ss.toast('PR25 R8 자동 이어실행을 시작했습니다.', 'PR25 R8', 8);
    return { ok: true, scheduled: true, state: state };
  } finally {
    lock.releaseLock();
  }
}

function continuePr25TrackingContinuationR8() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: false, busy: true };
  try {
    const state = pr25r8_getState_();
    if (!state || state.status !== 'running') {
      pr25r8_clearTriggers_();
      return { ok: false, reason: 'NO_RUNNING_STATE' };
    }

    pr25r8_assertHelpers_();
    const ss = SpreadsheetApp.openById(state.spreadsheetId);
    state.nextRunScheduled = false;

    try {
      if (state.phase === 'backfill') {
        const result = pr25r8_backfillTrackingPayment_(ss, state);
        state.sourceEligibleRows = Number(result.sourceEligibleRows || 0);
        state.detailRows = Number(result.detailRows || 0);
        state.matchedDetailRows = Number(result.matchedDetailRows || 0);
        state.ignoredSourceRows = Number(result.ignoredSourceRows || 0);
        state.unmatchedDetailRows = Number(result.unmatchedDetailRows || 0);
        state.paymentRows = Number(result.paymentRows || 0);
        state.blankPaymentRows = Number(result.blankPaymentRows || 0);
        state.phase = 'match';
        state.updatedAt = new Date().toISOString();
        pr25r8_saveState_(state);
        pr25r8_schedule_(state);
        pr25r8_writeStatus_(ss, state, '상세행 100% 트래킹 백필 완료; 카드 재매칭 예약');
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
        pr25r8_saveState_(state);
        pr25r8_clearTriggers_();
        pr25r8_writeStatus_(ss, state, 'PR25 트래킹 결제수단 카드 재매칭 완료');
        ss.toast('PR25 R8 카드 재매칭이 완료되었습니다.', 'PR25 R8', 8);
        return { ok: true, done: true };
      }

      throw new Error('알 수 없는 PR25 단계: ' + state.phase);
    } catch (e) {
      state.status = 'failed';
      state.nextRunScheduled = false;
      state.updatedAt = new Date().toISOString();
      state.lastError = String(e && e.message ? e.message : e);
      pr25r8_saveState_(state);
      pr25r8_clearTriggers_();
      pr25r8_writeStatus_(ss, state, '오류로 중단');
      try { ss.toast('PR25 R8 작업이 중단되었습니다. PR25_실행상태를 확인하세요.', 'PR25 R8', 10); } catch (ignore) {}
      throw e;
    }
  } finally {
    lock.releaseLock();
  }
}

function pr25r8_backfillTrackingPayment_(ss, state) {
  const source = ss.getSheetByName('매출데이터_붙여넣기');
  const detail = ss.getSheetByName('부가세_신고자료');
  if (!source || source.getLastRow() < 2) throw new Error('매출데이터_붙여넣기 데이터가 없습니다.');
  if (!detail || detail.getLastRow() < 2) throw new Error('부가세_신고자료 데이터가 없습니다.');

  const lastRow = source.getLastRow();
  if (source.getLastColumn() < PR25_R8_CORE_COLS) {
    throw new Error('원본 A:AC 범위가 부족합니다. 현재 열수: ' + source.getLastColumn());
  }

  const found = pr25r8_discoverPaymentColumns_(source);
  if (found.trackingIndex < 0 && found.fallbackIndex < 0) {
    throw new Error('원본 전체 열에서 트래킹 번호 또는 결제수단 열을 찾지 못했습니다.');
  }
  state.sourceLastColumn = found.lastColumn;
  state.trackingIndex = found.trackingIndex;
  state.trackingHeader = found.trackingHeader;
  state.fallbackIndex = found.fallbackIndex;
  state.fallbackHeader = found.fallbackHeader;
  pr25r8_saveState_(state);

  const sourceValues = source.getRange(1, 1, lastRow, PR25_R8_CORE_COLS).getValues();
  const sh = sourceValues[0] || [];
  const sx = {
    date: pr25r5_find_(sh, ['마켓주문일자','주문일자','결제일자','주문일시'], 0),
    account: 3,
    orderNo: pr25r5_find_(sh, ['마켓주문번호','주문번호','주문ID','주문ID(마켓)'], 2),
    customer: pr25r5_find_(sh, ['고객명','수령인','수취인','구매자','주문자'], -1),
    brand: pr25r5_find_(sh, ['브랜드명','브랜드'], -1),
    productNo: pr25r5_find_(sh, ['마켓상품번호','상품번호','상품코드','판매자상품코드'], 4),
    productName: pr25r5_find_(sh, ['상품명','상품명(옵션포함)','등록상품명'], -1),
    quantity: pr25r5_find_(sh, ['판매수량','수량','구매수량'], -1),
    sales: pr25r5_find_(sh, ['결제금액합계(원)','결제금액합계','결제금액','순수매출액','판매금액'], 6),
    status: pr25r5_find_(sh, ['주문상태','상태','클레임상태','처리상태'], -1),
    purchase: 28
  };

  const trackingValues = found.trackingIndex >= 0
    ? source.getRange(2, found.trackingIndex + 1, lastRow - 1, 1).getDisplayValues()
    : null;
  const fallbackValues = found.fallbackIndex >= 0
    ? source.getRange(2, found.fallbackIndex + 1, lastRow - 1, 1).getDisplayValues()
    : null;

  const detailValues = detail.getDataRange().getValues();
  const dh = (detailValues[0] || []).slice();
  const existingPaymentIx = pr25r5_find_(dh, ['롯데결제수단'], -1);
  const paymentIx = existingPaymentIx >= 0 ? existingPaymentIx : dh.length;
  if (existingPaymentIx < 0) dh.push('롯데결제수단');

  const dx = {
    date: pr25r5_find_(dh, ['날짜','주문일','주문일자','마켓주문일자'], -1),
    account: pr25r5_find_(dh, ['쿠팡계정ID'], -1),
    orderNo: pr25r5_find_(dh, ['주문번호','마켓주문번호','주문ID','주문ID(마켓)'], -1),
    customer: pr25r5_find_(dh, ['고객명'], -1),
    brand: pr25r5_find_(dh, ['브랜드명'], -1),
    productNo: pr25r5_find_(dh, ['상품번호','마켓상품번호'], -1),
    productName: pr25r5_find_(dh, ['상품명'], -1),
    quantity: pr25r5_find_(dh, ['판매수량','수량'], -1),
    sales: pr25r5_find_(dh, ['순수매출액'], -1),
    purchase: pr25r5_find_(dh, ['매입금액'], -1)
  };
  Object.keys(dx).forEach(function(key) {
    if (dx[key] < 0) throw new Error('부가세_신고자료 필수 열 누락: ' + key);
  });

  const queues = {};
  let sourceEligibleRows = 0;
  for (let r = 1; r < sourceValues.length; r++) {
    const row = sourceValues[r];
    const status = pr25r5_text_(pr25r5_at_(row, sx.status));
    if (/취소|반품|환불/.test(status)) continue;
    const sales = pr25r5_num_(pr25r5_at_(row, sx.sales));
    if (!sales) continue;

    const generatedRow = {
      date: pr25r5_date_(pr25r5_at_(row, sx.date)),
      account: pr25r5_text_(row[sx.account]),
      orderNo: pr25r5_text_(pr25r5_at_(row, sx.orderNo)),
      customer: pr25r5_text_(pr25r5_at_(row, sx.customer)),
      brand: pr25r5_text_(pr25r5_at_(row, sx.brand)),
      productNo: pr25r5_text_(pr25r5_at_(row, sx.productNo)),
      productName: pr25r5_text_(pr25r5_at_(row, sx.productName)),
      quantity: pr25r5_num_(pr25r5_at_(row, sx.quantity)) || 1,
      sales: sales,
      purchase: pr25r5_num_(row[sx.purchase])
    };

    const offset = r - 1;
    const tracking = trackingValues ? pr25r5_text_(trackingValues[offset][0]) : '';
    const fallback = fallbackValues ? pr25r5_text_(fallbackValues[offset][0]) : '';
    const payment = tracking || fallback;
    const rowKey = pr25r5_rowKey_(generatedRow);
    if (!queues[rowKey]) queues[rowKey] = [];
    queues[rowKey].push({ payment: payment, sourceRow: r + 1 });
    sourceEligibleRows++;
  }

  const detailRows = detailValues.slice(1);
  const payments = [];
  const unmatched = [];
  let paymentRows = 0;
  let matchedDetailRows = 0;

  for (let i = 0; i < detailRows.length; i++) {
    const row = detailRows[i];
    const obj = {
      date: pr25r5_date_(row[dx.date]),
      account: pr25r5_text_(row[dx.account]),
      orderNo: pr25r5_text_(row[dx.orderNo]),
      customer: pr25r5_text_(row[dx.customer]),
      brand: pr25r5_text_(row[dx.brand]),
      productNo: pr25r5_text_(row[dx.productNo]),
      productName: pr25r5_text_(row[dx.productName]),
      quantity: pr25r5_num_(row[dx.quantity]) || 1,
      sales: pr25r5_num_(row[dx.sales]),
      purchase: pr25r5_num_(row[dx.purchase])
    };
    const rowKey = pr25r5_rowKey_(obj);
    const queue = queues[rowKey];
    if (!queue || !queue.length) {
      unmatched.push(i + 2);
      payments.push(['']);
      continue;
    }
    const candidate = queue.shift();
    payments.push([candidate.payment]);
    matchedDetailRows++;
    if (candidate.payment) paymentRows++;
  }

  if (unmatched.length) {
    const sample = unmatched.slice(0, 20).join(', ');
    throw new Error(
      '안전검증 실패: 상세 미매칭 ' + unmatched.length + '건' +
      (sample ? ' (예시 행: ' + sample + ')' : '') + '. 아무 값도 쓰지 않았습니다.'
    );
  }
  if (matchedDetailRows !== detailRows.length) {
    throw new Error('안전검증 실패: 상세행 ' + detailRows.length + '건 중 ' + matchedDetailRows + '건만 매칭되었습니다. 아무 값도 쓰지 않았습니다.');
  }
  if (paymentRows < 1) {
    throw new Error('매칭된 상세행의 결제수단 값이 0건입니다. 아무 값도 쓰지 않았습니다.');
  }

  let ignoredSourceRows = 0;
  Object.keys(queues).forEach(function(key) { ignoredSourceRows += queues[key].length; });

  if (existingPaymentIx < 0) detail.getRange(1, paymentIx + 1).setValue('롯데결제수단');
  detail.getRange(2, paymentIx + 1, payments.length, 1).setNumberFormat('@').setValues(payments);
  SpreadsheetApp.flush();

  return {
    sourceEligibleRows: sourceEligibleRows,
    detailRows: detailRows.length,
    matchedDetailRows: matchedDetailRows,
    ignoredSourceRows: ignoredSourceRows,
    unmatchedDetailRows: 0,
    paymentRows: paymentRows,
    blankPaymentRows: detailRows.length - paymentRows
  };
}

function pr25r8_discoverPaymentColumns_(source) {
  const lastColumn = source.getLastColumn();
  const headers = source.getRange(1, 1, 1, lastColumn).getDisplayValues()[0] || [];
  let trackingIndex = pr25r8_findHeader_(headers, ['트래킹 번호','트래킹번호','tracking number','trackingnumber']);
  let fallbackIndex = pr25r8_findHeader_(headers, ['결제수단','결제정보','결제방법','카드사','결제수단/카드사','결제수단(카드사)','구매결제수단']);

  if (trackingIndex < 0) {
    for (let i = 0; i < headers.length; i++) {
      const compact = pr25r8_compactHeader_(headers[i]);
      if (compact.indexOf('트래킹번호') >= 0 || compact.indexOf('trackingnumber') >= 0) {
        trackingIndex = i;
        break;
      }
    }
  }
  if (fallbackIndex < 0) {
    for (let i = 0; i < headers.length; i++) {
      const compact = pr25r8_compactHeader_(headers[i]);
      if (compact.indexOf('결제수단') >= 0 || compact.indexOf('결제정보') >= 0 || compact.indexOf('결제방법') >= 0) {
        fallbackIndex = i;
        break;
      }
    }
  }

  const relatedHeaders = [];
  headers.forEach(function(header, index) {
    const compact = pr25r8_compactHeader_(header);
    if (/트래킹|tracking|결제|카드|페이/.test(compact)) relatedHeaders.push((index + 1) + '열:' + String(header || ''));
  });

  return {
    lastColumn: lastColumn,
    trackingIndex: trackingIndex,
    trackingHeader: trackingIndex >= 0 ? String(headers[trackingIndex] || '') : '',
    fallbackIndex: fallbackIndex,
    fallbackHeader: fallbackIndex >= 0 ? String(headers[fallbackIndex] || '') : '',
    relatedHeaders: relatedHeaders
  };
}

function pr25r8_findHeader_(headers, names) {
  for (let n = 0; n < names.length; n++) {
    const target = pr25r8_compactHeader_(names[n]);
    for (let i = 0; i < headers.length; i++) {
      if (pr25r8_compactHeader_(headers[i]) === target) return i;
    }
  }
  return -1;
}

function pr25r8_compactHeader_(value) {
  return String(value == null ? '' : value).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g, '');
}

function resetPr25TrackingContinuationR8() {
  pr25r8_clearTriggers_();
  PropertiesService.getScriptProperties().deleteProperty(PR25_R8_STATE_KEY);
  const ss = SpreadsheetApp.getActive();
  pr25r8_writeStatus_(ss, {
    version: PR25_R8_VERSION,
    status: 'reset',
    phase: '-',
    nextRunScheduled: false,
    startedAt: '',
    updatedAt: new Date().toISOString(),
    lastError: ''
  }, '상태 초기화 완료');
}

function pr25r8_assertHelpers_() {
  const required = [
    'pr25r5_find_','pr25r5_at_','pr25r5_text_','pr25r5_num_','pr25r5_date_','pr25r5_rowKey_',
    'pr25r5_groupOrders_','pr25r5_loadHistory_','pr25r5_loadMaster_','pr25r5_canonicalize_',
    'pr25r5_allocate_','pr25r5_writeDiagnostic_','pr25r5_writeSummary_','pr25r5_writeMonthlyPayment_'
  ];
  const missing = required.filter(function(name) {
    try { return typeof eval(name) !== 'function'; } catch (e) { return true; }
  });
  if (missing.length) throw new Error('PR25 R5 helper 누락: ' + missing.join(', '));
}

function pr25r8_assertSheets_(ss) {
  const required = ['매출데이터_붙여넣기','부가세_신고자료','카드사용내역_붙여넣기'];
  const missing = required.filter(function(name) { return !ss.getSheetByName(name); });
  if (missing.length) throw new Error('필수 시트 누락: ' + missing.join(', '));
}

function pr25r8_getState_() {
  const raw = PropertiesService.getScriptProperties().getProperty(PR25_R8_STATE_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function pr25r8_saveState_(state) {
  PropertiesService.getScriptProperties().setProperty(PR25_R8_STATE_KEY, JSON.stringify(state));
}

function pr25r8_schedule_(state) {
  pr25r8_clearTriggers_();
  ScriptApp.newTrigger(PR25_R8_HANDLER).timeBased().after(PR25_R8_DELAY_MS).create();
  state.nextRunScheduled = true;
  state.updatedAt = new Date().toISOString();
  pr25r8_saveState_(state);
}

function pr25r8_clearTriggers_() {
  const oldHandlers = {
    continuePr25TrackingContinuationR6: true,
    continuePr25TrackingContinuationR7: true,
    continuePr25TrackingContinuationR8: true
  };
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    try {
      if (oldHandlers[trigger.getHandlerFunction()]) ScriptApp.deleteTrigger(trigger);
    } catch (e) {}
  });
}

function pr25r8_writeStatus_(ss, state, message) {
  const sheet = ss.getSheetByName(PR25_R8_STATUS_SHEET) || ss.insertSheet(PR25_R8_STATUS_SHEET);
  const trackingCol = Number(state.trackingIndex) >= 0 ? Number(state.trackingIndex) + 1 : '';
  const fallbackCol = Number(state.fallbackIndex) >= 0 ? Number(state.fallbackIndex) + 1 : '';
  const rows = [
    ['항목','값'],
    ['버전',state.version || PR25_R8_VERSION],
    ['상태',state.status || ''],
    ['단계',state.phase || ''],
    ['메시지',message || ''],
    ['다음 실행 예약',state.nextRunScheduled ? 'Y' : 'N'],
    ['시작시각',state.startedAt || ''],
    ['갱신시각',state.updatedAt || ''],
    ['원본 마지막 열',Number(state.sourceLastColumn || 0)],
    ['트래킹 헤더',state.trackingHeader || ''],
    ['트래킹 열번호',trackingCol],
    ['대체 결제수단 헤더',state.fallbackHeader || ''],
    ['대체 결제수단 열번호',fallbackCol],
    ['원본 유효행',Number(state.sourceEligibleRows || 0)],
    ['부가세 상세행',Number(state.detailRows || 0)],
    ['상세 매칭행',Number(state.matchedDetailRows || 0)],
    ['제외한 원본 추가행',Number(state.ignoredSourceRows || 0)],
    ['상세 미매칭행',Number(state.unmatchedDetailRows || 0)],
    ['결제수단 입력행',Number(state.paymentRows || 0)],
    ['결제수단 공란행',Number(state.blankPaymentRows || 0)],
    ['주문건수',Number(state.orderRows || 0)],
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
  sheet.setColumnWidth(2, 650);
  SpreadsheetApp.flush();
}
