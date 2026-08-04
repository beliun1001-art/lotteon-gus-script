/**
 * PR #39 / Issue #38 remote production apply.
 * Applies PR33 PASS preview to production VAT period/diagnostic sheets.
 * Uses permanent remote trigger bridge and 150-row automatic continuation.
 */
const PR38_VERSION = 'v1.0-PR38-V670-PRODUCTION-BATCHED';
const PR38_HANDLER = 'runLotteonRemoteTaskContinue';
const PR38_STATE_KEY = 'PR38_REMOTE_APPLY_STATE';
const PR38_BATCH = 150;

const PR38_PREVIEW_STATUS = 'PR33_실행상태';
const PR38_PREVIEW_DIAG = 'PR33_카드매칭검증';
const PR38_VAT_DETAIL = '부가세_신고자료';
const PR38_PROD_PERIOD = '부가세_기간별';
const PR38_PROD_DIAG = '부가세_카드매칭검증';
const PR38_BACKUP_PERIOD = 'PR38_백업_부가세_기간별';
const PR38_BACKUP_DIAG = 'PR38_백업_부가세_카드매칭검증';
const PR38_SUMMARY_TEMP = 'PR38_요약준비';
const PR38_STATUS = 'PR38_운영반영상태';

function runPr38ProductionApplyStart() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  pr38DeleteTriggers_();

  const previous = pr38StatusValue_(ss, '상태');
  const backups = !!ss.getSheetByName(PR38_BACKUP_PERIOD) && !!ss.getSheetByName(PR38_BACKUP_DIAG);
  const state = {
    spreadsheetId: ss.getId(),
    stage: backups && previous !== 'PASS' ? 'RECOVER_PERIOD' : 'VALIDATE',
    offset: 0,
    startedAt: new Date().toISOString(),
    completedAt: '',
    error: '',
    diagRows: 0,
    summaryRows: 0,
    periodSourceStart: 0,
    periodDestStart: 0,
    periodDetailRows: 0,
    periodDetailCols: 0,
    stats: pr38EmptyStats_()
  };
  pr38Save_(state);
  pr38WriteStatus_(ss, 'RUNNING', 'v6.70 운영 배치 반영 시작', state);
  return runPr38ProductionApplyContinue();
}

function runPr38ProductionApplyContinue() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return {ok:false, reason:'LOCK_BUSY'};
  try {
    const state = pr38Load_();
    if (!state || !state.spreadsheetId) throw new Error('PR38 실행 상태가 없습니다.');
    const ss = SpreadsheetApp.openById(state.spreadsheetId);
    const handlers = {
      RECOVER_PERIOD: pr38RecoverPeriod_,
      RECOVER_DIAG: pr38RecoverDiag_,
      VALIDATE: pr38Validate_,
      BACKUP_PERIOD: pr38BackupPeriod_,
      BACKUP_DIAG: pr38BackupDiag_,
      BUILD_SUMMARY: pr38BuildSummary_,
      PREP_PERIOD: pr38PrepPeriod_,
      WRITE_PERIOD: pr38WritePeriod_,
      PREP_DIAG: pr38PrepDiag_,
      WRITE_DIAG: pr38WriteDiag_,
      VERIFY_PERIOD_DETAIL: pr38VerifyPeriodDetail_,
      VERIFY_DIAG: pr38VerifyDiag_,
      ROLLBACK_PERIOD: pr38RollbackPeriod_,
      ROLLBACK_DIAG: pr38RollbackDiag_
    };
    const handler = handlers[state.stage];
    if (!handler) {
      if (state.stage === 'DONE') return {ok:true, done:true};
      throw new Error('알 수 없는 PR38 단계: ' + state.stage);
    }
    return handler(ss, state);
  } catch (error) {
    const state = pr38Load_() || {};
    let ss = null;
    try {
      ss = state.spreadsheetId ? SpreadsheetApp.openById(state.spreadsheetId) : SpreadsheetApp.getActive();
    } catch (ignore) {}

    state.error = String(error && error.message ? error.message : error);
    if (state.stage === 'ROLLBACK_PERIOD' || state.stage === 'ROLLBACK_DIAG') {
      state.stage = 'DONE';
      state.completedAt = new Date().toISOString();
      pr38Save_(state);
      if (ss) pr38WriteStatus_(ss, 'ROLLBACK_ERROR', state.error, state);
      pr38DeleteTriggers_();
      return {ok:false, rollbackError:true, error:state.error};
    }

    state.stage = ss && ss.getSheetByName(PR38_BACKUP_PERIOD) ? 'ROLLBACK_PERIOD' : 'DONE';
    pr38Save_(state);
    if (ss) pr38WriteStatus_(ss, state.stage === 'DONE' ? 'ERROR' : 'ROLLBACK_PENDING', state.error, state);
    if (state.stage !== 'DONE') pr38Schedule_();
    return {ok:false, error:state.error, rollbackScheduled:state.stage !== 'DONE'};
  } finally {
    lock.releaseLock();
  }
}

function pr38RecoverPeriod_(ss, state) {
  pr38RestoreBackup_(ss, PR38_BACKUP_PERIOD, PR38_PROD_PERIOD);
  return pr38Next_(ss, state, 'RECOVER_DIAG', '이전 미완료 기간별 백업 복구 완료');
}

function pr38RecoverDiag_(ss, state) {
  pr38RestoreBackup_(ss, PR38_BACKUP_DIAG, PR38_PROD_DIAG);
  return pr38Next_(ss, state, 'VALIDATE', '이전 미완료 카드검증 백업 복구 완료');
}

function pr38Validate_(ss, state) {
  ['groupVatDetailByOrder_v660_','aggregateVatBusinessCardHalf_v660_','vatBusinessCardHalfHeaders_v660_'].forEach(function(name) {
    if (typeof eval(name) !== 'function') throw new Error('main v6.70 필수 함수 없음: ' + name);
  });

  const requiredSheets = [PR38_PREVIEW_STATUS, PR38_PREVIEW_DIAG, PR38_VAT_DETAIL, PR38_PROD_PERIOD, PR38_PROD_DIAG];
  requiredSheets.forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 1) throw new Error('필수 시트 없음: ' + name);
  });

  const status = pr38KeyValue_(ss.getSheetByName(PR38_PREVIEW_STATUS).getDataRange().getValues());
  const expected = {
    '버전':'v1.1-PR33-V670-DIAG-BATCHED-LOCAL-RULE',
    '상태':'PASS',
    '운영시트 변경':'없음',
    '상반기 주문':1355,
    'MATCHED':810,
    'NON_CARD':494,
    'AMBIGUOUS':1,
    'NO_MATCH':50,
    'v6.70 3차귀속':44,
    '주문매입금액':54807644,
    '잘못된 카드 식별자':0,
    '3차귀속 증빙필드 오류':0
  };
  Object.keys(expected).forEach(function(key) {
    const actual = status[key];
    const want = expected[key];
    if (typeof want === 'number') {
      if (Math.round(pr38Number_(actual)) !== want) throw new Error('PR33 상태값 불일치: ' + key + ' 실제 ' + actual + ' / 기대 ' + want);
    } else if (pr38Text_(actual) !== String(want)) {
      throw new Error('PR33 상태값 불일치: ' + key + ' 실제 ' + actual + ' / 기대 ' + want);
    }
  });

  const preview = ss.getSheetByName(PR38_PREVIEW_DIAG);
  const headers = preview.getRange(1,1,1,preview.getLastColumn()).getValues()[0].map(pr38Text_);
  pr38Require_(headers, [
    '신고연도','반기','주문일','사업자등록번호','쿠팡계정ID','주문번호','롯데결제수단',
    '주문매입금액','구매카드사','구매카드명','카드번호','카드번호끝4',
    '승인일','승인번호','승인금액','카드매칭상태','카드매칭근거','v6.69 2차귀속','v6.70 3차귀속'
  ]);
  state.diagRows = preview.getLastRow() - 1;
  if (state.diagRows !== 1355) throw new Error('PR33 검증표 행 수 불일치: ' + state.diagRows);

  const stats = pr38CountDiag_(headers, preview.getRange(2,1,state.diagRows,headers.length).getValues());
  pr38AssertStats_(stats);
  state.stats = stats;
  state.offset = 0;
  return pr38Next_(ss, state, 'BACKUP_PERIOD', 'PR33 PASS 및 운영 원본 검증 완료');
}

function pr38BackupPeriod_(ss, state) {
  pr38ReplaceBackup_(ss, PR38_PROD_PERIOD, PR38_BACKUP_PERIOD);
  return pr38Next_(ss, state, 'BACKUP_DIAG', '부가세_기간별 백업 완료');
}

function pr38BackupDiag_(ss, state) {
  pr38ReplaceBackup_(ss, PR38_PROD_DIAG, PR38_BACKUP_DIAG);
  return pr38Next_(ss, state, 'BUILD_SUMMARY', '부가세_카드매칭검증 백업 완료');
}

function pr38BuildSummary_(ss, state) {
  const detailSheet = ss.getSheetByName(PR38_VAT_DETAIL);
  const detailValues = detailSheet.getDataRange().getValues();
  let orders = groupVatDetailByOrder_v660_(detailValues).filter(function(order) {
    return String(order.year) === '2026' && String(order.half) === '상반기';
  });
  if (orders.length !== 1355) throw new Error('상반기 주문 집계 불일치: ' + orders.length);

  const preview = ss.getSheetByName(PR38_PREVIEW_DIAG);
  const values = preview.getDataRange().getValues();
  const headers = values[0].map(pr38Text_);
  const lookup = pr38BuildDiagLookup_(headers, values.slice(1));

  let matched = 0;
  orders.forEach(function(order) {
    const match = pr38TakeDiagMatch_(lookup, order);
    if (!match) throw new Error('PR33 카드결과 연결 실패: ' + pr38OrderLabel_(order));
    order.cardMatch = match;
    matched++;
  });
  if (matched !== 1355 || lookup.remaining !== 0) {
    throw new Error('PR33 카드결과 연결 건수 불일치: matched=' + matched + ' / remaining=' + lookup.remaining);
  }

  const summary = aggregateVatBusinessCardHalf_v660_(orders);
  const summaryHeaders = vatBusinessCardHalfHeaders_v660_();
  const totals = pr38SummaryTotals_(summaryHeaders, summary);
  const expectedTotals = {
    orders:1355, sales:71838700, salesSupply:65307938, salesVat:6530762,
    settlement:64726771, fee:7111929, purchase:54807644,
    purchaseSupply:49825146, purchaseVat:4982498, payable:1548264,
    profit:9919127, vatProfit:8370863
  };
  Object.keys(expectedTotals).forEach(function(key) {
    if (Math.round(totals[key]) !== expectedTotals[key]) {
      throw new Error('구매카드별 요약 합계 불일치: ' + key + ' 실제 ' + totals[key] + ' / 기대 ' + expectedTotals[key]);
    }
  });

  const old = ss.getSheetByName(PR38_SUMMARY_TEMP);
  if (old) ss.deleteSheet(old);
  const temp = ss.insertSheet(PR38_SUMMARY_TEMP);
  temp.getRange(1,1,1,summaryHeaders.length).setValues([summaryHeaders]);
  if (summary.length) temp.getRange(2,1,summary.length,summaryHeaders.length).setValues(summary);
  temp.setFrozenRows(1);
  temp.getRange(1,1,1,summaryHeaders.length).setBackground('#d9eaf7').setFontWeight('bold');
  state.summaryRows = summary.length;
  state.offset = 0;
  return pr38Next_(ss, state, 'PREP_PERIOD', '구매카드별 반기 요약 재집계 완료');
}

function pr38PrepPeriod_(ss, state) {
  const backup = ss.getSheetByName(PR38_BACKUP_PERIOD);
  const period = ss.getSheetByName(PR38_PROD_PERIOD);
  const summary = ss.getSheetByName(PR38_SUMMARY_TEMP);
  if (!backup || !period || !summary) throw new Error('기간별 반영 준비 시트가 없습니다.');

  const detail = pr38FindPeriodDetail_(backup);
  state.periodSourceStart = detail.startRow;
  state.periodDetailRows = detail.rows;
  state.periodDetailCols = detail.cols;
  state.periodDestStart = state.summaryRows + 5;
  state.offset = 0;

  const summaryValues = summary.getDataRange().getValues();
  const headers = summaryValues[0].map(pr38Text_);
  const rows = summaryValues.slice(1);

  period.clear();
  period.getRange(1,1).setValue('사업자별 반기 신고요약 (구매카드별)');
  period.getRange(2,1,1,headers.length).setValues([headers]);
  if (rows.length) period.getRange(3,1,rows.length,headers.length).setValues(rows);
  period.getRange(1,1,1,Math.max(headers.length,state.periodDetailCols)).setBackground('#b4c6e7').setFontWeight('bold');
  period.getRange(2,1,1,headers.length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
  if (rows.length) {
    period.getRange(3,1,rows.length,11).setNumberFormat('@');
    for (let c=12;c<=23;c++) period.getRange(3,c,rows.length,1).setNumberFormat('#,##0');
  }
  period.setFrozenRows(2);
  return pr38Next_(ss, state, 'WRITE_PERIOD', '부가세_기간별 배치 쓰기 준비 완료');
}

function pr38WritePeriod_(ss, state) {
  const backup = ss.getSheetByName(PR38_BACKUP_PERIOD);
  const period = ss.getSheetByName(PR38_PROD_PERIOD);
  const count = Math.min(PR38_BATCH, state.periodDetailRows - state.offset);
  if (count > 0) {
    const src = backup.getRange(state.periodSourceStart + state.offset, 1, count, state.periodDetailCols);
    const dst = period.getRange(state.periodDestStart + state.offset, 1, count, state.periodDetailCols);
    src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    state.offset += count;
  }
  if (state.offset >= state.periodDetailRows) {
    state.offset = 0;
    return pr38Next_(ss, state, 'PREP_DIAG', '부가세_기간별 상세표 배치 반영 완료');
  }
  return pr38Stay_(ss, state, '기간별 상세표 배치 반영 ' + state.offset + ' / ' + state.periodDetailRows);
}

function pr38PrepDiag_(ss, state) {
  const src = ss.getSheetByName(PR38_PREVIEW_DIAG);
  const dst = ss.getSheetByName(PR38_PROD_DIAG);
  const headers = src.getRange(1,1,1,src.getLastColumn()).getValues()[0].map(pr38Text_);
  dst.clear();
  src.getRange(1,1,1,headers.length).copyTo(dst.getRange(1,1,1,headers.length), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  dst.setFrozenRows(1);
  headers.forEach(function(header, index) {
    dst.setColumnWidth(index+1, /후보요약|취소|근거|원본파일/.test(header) ? 220 : (/카드|사업자|계정|주문번호/.test(header) ? 135 : (/금액/.test(header) ? 105 : 90)));
  });
  state.offset = 0;
  return pr38Next_(ss, state, 'WRITE_DIAG', '카드매칭검증 배치 쓰기 준비 완료');
}

function pr38WriteDiag_(ss, state) {
  const src = ss.getSheetByName(PR38_PREVIEW_DIAG);
  const dst = ss.getSheetByName(PR38_PROD_DIAG);
  const cols = src.getLastColumn();
  const count = Math.min(PR38_BATCH, state.diagRows - state.offset);
  if (count > 0) {
    src.getRange(state.offset+2,1,count,cols).copyTo(
      dst.getRange(state.offset+2,1,count,cols),
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false
    );
    state.offset += count;
  }
  if (state.offset >= state.diagRows) {
    state.offset = 0;
    return pr38Next_(ss, state, 'VERIFY_PERIOD_DETAIL', '카드매칭검증 1,355건 배치 반영 완료');
  }
  return pr38Stay_(ss, state, '카드매칭검증 배치 반영 ' + state.offset + ' / ' + state.diagRows);
}

function pr38VerifyPeriodDetail_(ss, state) {
  const backup = ss.getSheetByName(PR38_BACKUP_PERIOD);
  const period = ss.getSheetByName(PR38_PROD_PERIOD);
  if (state.offset === 0) pr38VerifySummary_(ss, state);

  const count = Math.min(PR38_BATCH, state.periodDetailRows - state.offset);
  if (count > 0) {
    const expected = backup.getRange(state.periodSourceStart + state.offset,1,count,state.periodDetailCols).getValues();
    const actual = period.getRange(state.periodDestStart + state.offset,1,count,state.periodDetailCols).getValues();
    pr38AssertMatrix_('기간별 상세표 배치 ' + state.offset, expected, actual);
    state.offset += count;
  }
  if (state.offset >= state.periodDetailRows) {
    state.offset = 0;
    state.stats = pr38EmptyStats_();
    return pr38Next_(ss, state, 'VERIFY_DIAG', '구매카드 요약 및 기간 상세표 검증 완료');
  }
  return pr38Stay_(ss, state, '기간별 상세표 배치 검증 ' + state.offset + ' / ' + state.periodDetailRows);
}

function pr38VerifyDiag_(ss, state) {
  const src = ss.getSheetByName(PR38_PREVIEW_DIAG);
  const dst = ss.getSheetByName(PR38_PROD_DIAG);
  const headers = src.getRange(1,1,1,src.getLastColumn()).getValues()[0].map(pr38Text_);
  const count = Math.min(PR38_BATCH, state.diagRows - state.offset);
  if (count > 0) {
    const expected = src.getRange(state.offset+2,1,count,headers.length).getValues();
    const actual = dst.getRange(state.offset+2,1,count,headers.length).getValues();
    pr38AssertMatrix_('카드매칭검증 배치 ' + state.offset, expected, actual);
    pr38CountDiagInto_(headers, actual, state.stats);
    state.offset += count;
  }
  if (state.offset < state.diagRows) {
    return pr38Stay_(ss, state, '카드매칭검증 배치 검증 ' + state.offset + ' / ' + state.diagRows);
  }

  pr38AssertStats_(state.stats);
  state.stage = 'DONE';
  state.completedAt = new Date().toISOString();
  pr38Save_(state);
  pr38DeleteTriggers_();
  pr38WriteStatus_(ss, 'PASS', 'v6.70 운영 배치 반영 및 검증 완료', state);
  PropertiesService.getScriptProperties().deleteProperty(PR38_STATE_KEY);
  ss.toast('PR38 완료: MATCHED 810 / NO_MATCH 50', 'LOTTEON 자동화', 10);
  return {ok:true, done:true, stats:state.stats};
}

function pr38RollbackPeriod_(ss, state) {
  pr38RestoreBackup_(ss, PR38_BACKUP_PERIOD, PR38_PROD_PERIOD);
  return pr38Next_(ss, state, 'ROLLBACK_DIAG', '부가세_기간별 자동 롤백 완료', 'ROLLBACK_RUNNING');
}

function pr38RollbackDiag_(ss, state) {
  pr38RestoreBackup_(ss, PR38_BACKUP_DIAG, PR38_PROD_DIAG);
  state.stage = 'DONE';
  state.completedAt = new Date().toISOString();
  pr38Save_(state);
  pr38DeleteTriggers_();
  pr38WriteStatus_(ss, 'ROLLED_BACK', state.error || '오류로 자동 롤백', state);
  PropertiesService.getScriptProperties().deleteProperty(PR38_STATE_KEY);
  return {ok:false, rolledBack:true, error:state.error || ''};
}

function pr38BuildDiagLookup_(headers, rows) {
  const ix = pr38HeaderIndex_(headers);
  const map = {};
  let remaining = 0;
  rows.forEach(function(row) {
    const item = {
      year:pr38Text_(row[ix['신고연도']]),
      half:pr38Text_(row[ix['반기']]),
      orderDate:pr38DateKey_(row[ix['주문일']]),
      business:pr38Text_(row[ix['사업자등록번호']]),
      account:pr38Text_(row[ix['쿠팡계정ID']]),
      orderNo:pr38Text_(row[ix['주문번호']]),
      payment:pr38Text_(row[ix['롯데결제수단']]),
      purchase:pr38Number_(row[ix['주문매입금액']]),
      match:{
        company:pr38Text_(row[ix['구매카드사']]),
        alias:ix['구매카드별칭'] >= 0 ? pr38Text_(row[ix['구매카드별칭']]) : '',
        cardName:pr38Text_(row[ix['구매카드명']]),
        cardNumber:pr38Text_(row[ix['카드번호']]),
        cardEnd4:pr38Text_(row[ix['카드번호끝4']]),
        approvalDate:ix['승인일'] >= 0 ? row[ix['승인일']] : '',
        approvalTime:ix['승인시각'] >= 0 ? row[ix['승인시각']] : '',
        approvalNo:ix['승인번호'] >= 0 ? pr38Text_(row[ix['승인번호']]) : '',
        approvalAmount:ix['승인금액'] >= 0 ? pr38Number_(row[ix['승인금액']]) : 0,
        status:pr38Text_(row[ix['카드매칭상태']]),
        reason:pr38Text_(row[ix['카드매칭근거']]),
        candidateCount:ix['후보수'] >= 0 ? pr38Number_(row[ix['후보수']]) : 0,
        merchant:ix['가맹점명'] >= 0 ? pr38Text_(row[ix['가맹점명']]) : '',
        merchantOrderNo:ix['가맹점주문번호'] >= 0 ? pr38Text_(row[ix['가맹점주문번호']]) : '',
        evidenceType:ix['증빙유형'] >= 0 ? pr38Text_(row[ix['증빙유형']]) : '',
        cancelMemo:ix['취소/부분취소메모'] >= 0 ? pr38Text_(row[ix['취소/부분취소메모']]) : '',
        sourceFile:ix['원본파일'] >= 0 ? pr38Text_(row[ix['원본파일']]) : '',
        candidateSummary:ix['후보요약'] >= 0 ? pr38Text_(row[ix['후보요약']]) : ''
      }
    };
    const key = pr38DiagItemKey_(item);
    if (!map[key]) map[key] = [];
    map[key].push(item.match);
    remaining++;
  });
  return {map:map, remaining:remaining};
}

function pr38TakeDiagMatch_(lookup, order) {
  const item = {
    year:order.year, half:order.half, orderDate:order.orderDate, business:order.business,
    account:order.account, orderNo:order.orderNo, payment:order.lottePayment, purchase:order.purchase
  };
  const key = pr38DiagItemKey_(item);
  const queue = lookup.map[key];
  if (!queue || !queue.length) return null;
  lookup.remaining--;
  return queue.shift();
}

function pr38DiagItemKey_(item) {
  const base = [pr38Text_(item.year),pr38Text_(item.half),pr38Text_(item.business),pr38Text_(item.account)];
  const orderNo = pr38Text_(item.orderNo);
  if (orderNo) return base.concat(['ORDER',orderNo]).join('|');
  return base.concat([
    'BLANK',pr38DateKey_(item.orderDate),Math.round(pr38Number_(item.purchase)),pr38Compact_(item.payment)
  ]).join('|');
}

function pr38OrderLabel_(order) {
  return [order.year,order.half,order.business,order.account,order.orderNo || '(공란)',order.orderDate,order.purchase].join(' / ');
}

function pr38SummaryTotals_(headers, rows) {
  const ix = pr38HeaderIndex_(headers);
  const out = {orders:0,sales:0,salesSupply:0,salesVat:0,settlement:0,fee:0,purchase:0,purchaseSupply:0,purchaseVat:0,payable:0,profit:0,vatProfit:0};
  const fields = {
    orders:'주문건수',sales:'순수매출액',salesSupply:'매출공급가액',salesVat:'매출부가세',
    settlement:'정산기준금액',fee:'마켓수수료',purchase:'매입금액',purchaseSupply:'매입공급가액',
    purchaseVat:'매입부가세',payable:'납부예상부가세',profit:'예상이익',vatProfit:'부가세반영예상이익'
  };
  rows.forEach(function(row) {
    Object.keys(fields).forEach(function(key) { out[key] += pr38Number_(row[ix[fields[key]]]); });
  });
  return out;
}

function pr38VerifySummary_(ss, state) {
  const temp = ss.getSheetByName(PR38_SUMMARY_TEMP);
  const period = ss.getSheetByName(PR38_PROD_PERIOD);
  const expected = temp.getDataRange().getValues();
  const actual = period.getRange(2,1,expected.length,expected[0].length).getValues();
  pr38AssertMatrix_('구매카드별 반기 요약', expected, actual);
}

function pr38FindPeriodDetail_(sheet) {
  const values = sheet.getDataRange().getValues();
  const index = values.findIndex(function(row, i) {
    return i >= 2 && pr38Text_(row[0]) === '집계구분' && row.map(pr38Text_).indexOf('신고연도') >= 0;
  });
  if (index < 0) throw new Error('부가세_기간별 상세표 시작행을 찾지 못했습니다.');
  let cols = values[index].length;
  while (cols > 1 && pr38Text_(values[index][cols-1]) === '') cols--;
  return {startRow:index+1, rows:values.length-index, cols:Math.max(1,cols)};
}

function pr38CountDiag_(headers, rows) {
  const stats = pr38EmptyStats_();
  pr38CountDiagInto_(headers, rows, stats);
  return stats;
}

function pr38CountDiagInto_(headers, rows, stats) {
  const ix = pr38HeaderIndex_(headers);
  rows.forEach(function(row) {
    const status = pr38Text_(row[ix['카드매칭상태']]);
    if (status === 'MATCHED' || status === 'MASTER_MATCHED') stats.matched++;
    else if (status === 'NON_CARD') stats.nonCard++;
    else if (status === 'AMBIGUOUS') stats.ambiguous++;
    else stats.noMatch++;
    if (ix['v6.69 2차귀속'] >= 0 && pr38Text_(row[ix['v6.69 2차귀속']]) === 'Y') stats.v669++;
    if (ix['v6.70 3차귀속'] >= 0 && pr38Text_(row[ix['v6.70 3차귀속']]) === 'Y') stats.v670++;
    stats.purchase += pr38Number_(row[ix['주문매입금액']]);
    stats.orders++;
  });
}

function pr38AssertStats_(stats) {
  const expected = {orders:1355,matched:810,nonCard:494,ambiguous:1,noMatch:50,v669:593,v670:44,purchase:54807644};
  Object.keys(expected).forEach(function(key) {
    if (Math.round(Number(stats[key] || 0)) !== expected[key]) {
      throw new Error('최종 상태 건수 불일치: ' + key + ' 실제 ' + stats[key] + ' / 기대 ' + expected[key]);
    }
  });
}

function pr38EmptyStats_() {
  return {orders:0,matched:0,nonCard:0,ambiguous:0,noMatch:0,v669:0,v670:0,purchase:0};
}

function pr38ReplaceBackup_(ss, sourceName, backupName) {
  const old = ss.getSheetByName(backupName);
  if (old) ss.deleteSheet(old);
  const source = ss.getSheetByName(sourceName);
  if (!source) throw new Error('백업 대상 시트 없음: ' + sourceName);
  source.copyTo(ss).setName(backupName);
}

function pr38RestoreBackup_(ss, backupName, productionName) {
  const backup = ss.getSheetByName(backupName);
  if (!backup) throw new Error('복구 백업 시트 없음: ' + backupName);
  const current = ss.getSheetByName(productionName);
  if (current) ss.deleteSheet(current);
  backup.copyTo(ss).setName(productionName);
}

function pr38Next_(ss, state, next, message, status) {
  state.stage = next;
  pr38Save_(state);
  pr38WriteStatus_(ss, status || 'RUNNING', message, state);
  pr38Schedule_();
  return {ok:true, stage:next};
}

function pr38Stay_(ss, state, message) {
  pr38Save_(state);
  pr38WriteStatus_(ss, 'RUNNING', message, state);
  pr38Schedule_();
  return {ok:true, stage:state.stage, offset:state.offset};
}

function pr38Schedule_() {
  pr38DeleteTriggers_();
  ScriptApp.newTrigger(PR38_HANDLER).timeBased().after(60*1000).create();
}

function pr38DeleteTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === PR38_HANDLER) {
      try { ScriptApp.deleteTrigger(trigger); } catch (ignore) {}
    }
  });
}

function pr38Save_(state) {
  PropertiesService.getScriptProperties().setProperty(PR38_STATE_KEY, JSON.stringify(state));
}

function pr38Load_() {
  const raw = PropertiesService.getScriptProperties().getProperty(PR38_STATE_KEY);
  return raw ? JSON.parse(raw) : null;
}

function pr38WriteStatus_(ss, status, message, state) {
  const stats = state.stats || pr38EmptyStats_();
  const sheet = ss.getSheetByName(PR38_STATUS) || ss.insertSheet(PR38_STATUS);
  const rows = [
    ['항목','값'],['버전',PR38_VERSION],['상태',status],['단계',state.stage || ''],
    ['메시지',message || ''],['처리행',state.offset || 0],['대상행',state.diagRows || 0],
    ['MATCHED',stats.matched || 0],['NON_CARD',stats.nonCard || 0],['AMBIGUOUS',stats.ambiguous || 0],
    ['NO_MATCH',stats.noMatch || 0],['2차귀속',stats.v669 || 0],['3차귀속',stats.v670 || 0],
    ['주문매입금액',stats.purchase || 0],['오류',state.error || ''],
    ['시작시각',state.startedAt || ''],['완료시각',state.completedAt || ''],['갱신시각',new Date().toISOString()]
  ];
  sheet.clearContents();
  sheet.getRange(1,1,rows.length,2).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');
  sheet.setColumnWidth(1,220);
  sheet.setColumnWidth(2,620);
}

function pr38StatusValue_(ss, key) {
  const sheet = ss.getSheetByName(PR38_STATUS);
  if (!sheet || sheet.getLastRow() < 2) return '';
  const values = sheet.getDataRange().getValues();
  for (let i=1;i<values.length;i++) if (pr38Text_(values[i][0]) === key) return pr38Text_(values[i][1]);
  return '';
}

function pr38AssertMatrix_(label, expected, actual) {
  const left = JSON.stringify(pr38Canonical_(expected));
  const right = JSON.stringify(pr38Canonical_(actual));
  if (left !== right) throw new Error(label + ' 불일치');
}

function pr38Canonical_(matrix) {
  return (matrix || []).map(function(row) {
    return (row || []).map(function(value) {
      if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return 'DATE:' + value.toISOString();
      if (typeof value === 'number') return 'NUM:' + String(value);
      if (typeof value === 'boolean') return 'BOOL:' + String(value);
      return 'TEXT:' + pr38Text_(value);
    });
  });
}

function pr38KeyValue_(values) {
  const out = {};
  (values || []).forEach(function(row) {
    const key = pr38Text_(row[0]);
    if (key) out[key] = row[1];
  });
  return out;
}

function pr38HeaderIndex_(headers) {
  const out = {};
  (headers || []).forEach(function(header,index) { out[pr38Text_(header)] = index; });
  return out;
}

function pr38Require_(headers, required) {
  required.forEach(function(header) {
    if (headers.indexOf(header) < 0) throw new Error('필수 헤더 없음: ' + header);
  });
}

function pr38Text_(value) {
  return String(value == null ? '' : value).trim();
}

function pr38Compact_(value) {
  return pr38Text_(value).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');
}

function pr38Number_(value) {
  const number = Number(typeof value === 'number' ? value : pr38Text_(value).replace(/[,원\s]/g,''));
  return isNaN(number) ? 0 : number;
}

function pr38DateKey_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
  }
  const text = pr38Text_(value);
  const match = text.match(/(20\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (match) return match[1] + '-' + ('0'+match[2]).slice(-2) + '-' + ('0'+match[3]).slice(-2);
  const date = new Date(text);
  return isNaN(date.getTime()) ? '' : Utilities.formatDate(date, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
}
