/**
 * PR #30 production apply runner.
 *
 * Applies the already validated PR29 preview result to production VAT summary/diagnostic sheets.
 * Reads production detail rows without recalculating card matches.
 * Creates exact-value backup sheets before any production write.
 */
const PR30_APPLY_VERSION = 'v1.0-PR30-APPLY-PR29-PASS';
const PR30_PREVIEW_STATUS_SHEET = 'PR29_실행상태';
const PR30_PREVIEW_SUMMARY_SHEET = 'PR29_사업자별반기요약';
const PR30_PREVIEW_DIAG_SHEET = 'PR29_카드매칭검증';
const PR30_PRODUCTION_PERIOD_SHEET = '부가세_기간별';
const PR30_PRODUCTION_DIAG_SHEET = '부가세_카드매칭검증';
const PR30_BACKUP_PERIOD_SHEET = 'PR30_백업_부가세_기간별';
const PR30_BACKUP_DIAG_SHEET = 'PR30_백업_부가세_카드매칭검증';
const PR30_STATUS_SHEET = 'PR30_적용상태';

function runPr30ApplyPreviewToProduction() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('다른 작업이 실행 중입니다. 잠시 후 다시 실행하세요.');

  let periodBefore = null;
  let diagBefore = null;
  try {
    const preview = pr30ReadAndValidatePreview_(ss);
    const production = pr30ReadProduction_(ss);

    periodBefore = production.periodValues;
    diagBefore = production.diagValues;

    pr30CreateBackupSheets_(ss, production.periodSheet, production.diagSheet);
    pr30WriteApplyStatus_(ss, 'WRITING', '백업 완료; 운영 시트 반영 중', preview);

    pr30ApplyPeriodSummary_(production.periodSheet, preview.summaryHeaders, preview.summaryRows, production.periodDetailValues);
    pr30ApplyDiagnostic_(production.diagSheet, preview.diagValues);
    SpreadsheetApp.flush();

    const verified = pr30ValidateProductionAfterWrite_(ss, preview, production.periodDetailValues);
    pr30WriteApplyStatus_(ss, 'PASS', '운영 반영 및 사후 검증 완료', verified);

    SpreadsheetApp.getUi().alert(
      'PR #30 운영 반영 완료\n\n' +
      '버전: ' + PR30_APPLY_VERSION + '\n' +
      '상반기 주문: ' + verified.orders + '건\n' +
      'MATCHED: ' + verified.matched + '건\n' +
      'NON_CARD: ' + verified.nonCard + '건\n' +
      'AMBIGUOUS: ' + verified.ambiguous + '건\n' +
      'NO_MATCH: ' + verified.noMatch + '건\n' +
      '2차귀속: ' + verified.fallback + '건\n\n' +
      '백업 시트: ' + PR30_BACKUP_PERIOD_SHEET + ', ' + PR30_BACKUP_DIAG_SHEET
    );
    return verified;
  } catch (error) {
    try {
      if (periodBefore && periodBefore.length) pr30RestoreValues_(ss, PR30_PRODUCTION_PERIOD_SHEET, periodBefore);
      if (diagBefore && diagBefore.length) pr30RestoreValues_(ss, PR30_PRODUCTION_DIAG_SHEET, diagBefore);
      SpreadsheetApp.flush();
      pr30WriteApplyStatus_(ss, 'ROLLED_BACK', String(error && error.message ? error.message : error), null);
    } catch (rollbackError) {
      pr30WriteApplyStatus_(ss, 'ROLLBACK_ERROR', String(rollbackError && rollbackError.message ? rollbackError.message : rollbackError), null);
    }
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function pr30ReadAndValidatePreview_(ss) {
  const statusSheet = ss.getSheetByName(PR30_PREVIEW_STATUS_SHEET);
  const summarySheet = ss.getSheetByName(PR30_PREVIEW_SUMMARY_SHEET);
  const diagSheet = ss.getSheetByName(PR30_PREVIEW_DIAG_SHEET);
  if (!statusSheet || !summarySheet || !diagSheet) {
    throw new Error('PR29 미리보기 시트가 없습니다. PR29 PASS 결과를 먼저 완료해야 합니다.');
  }

  const status = pr30KeyValueMap_(statusSheet.getDataRange().getValues());
  const expectedStatus = {
    '버전':'v1.2-PR29-V669-H1-BATCHED',
    '상태':'PASS',
    '운영시트 변경':'없음',
    '검증 대상':'2026년 상반기',
    '전체 기간 주문':1398,
    '대상 제외 주문':43,
    '상반기 주문':1355,
    'MATCHED':766,
    'NON_CARD':494,
    'AMBIGUOUS':24,
    'NO_MATCH':71,
    'v6.69 2차귀속':593,
    '2차귀속 MATCHED':341,
    '2차귀속 NON_CARD':229,
    '2차귀속 AMBIGUOUS':23,
    'canonical 증빙행':1990,
    '순수매출액':71838700,
    '매입금액':54807644,
    '납부예상부가세':1548264,
    '잘못된 카드 식별자':0,
    '2차귀속 증빙필드 오류':0
  };
  Object.keys(expectedStatus).forEach(function(key) {
    const actual = status[key];
    const expected = expectedStatus[key];
    if (typeof expected === 'number') {
      if (Math.round(pr30Number_(actual)) !== expected) throw new Error('PR29 상태값 불일치: ' + key + ' 실제 ' + actual + ' / 기대 ' + expected);
    } else if (String(actual || '') !== String(expected)) {
      throw new Error('PR29 상태값 불일치: ' + key + ' 실제 ' + actual + ' / 기대 ' + expected);
    }
  });

  const summaryValues = summarySheet.getDataRange().getValues();
  const diagValues = diagSheet.getDataRange().getValues();
  if (summaryValues.length < 2 || diagValues.length !== 1356) throw new Error('PR29 미리보기 행 수가 예상과 다릅니다.');

  const summaryHeaders = summaryValues[0].map(pr30Text_);
  const summaryRows = summaryValues.slice(1);
  const diagHeaders = diagValues[0].map(pr30Text_);
  const diagRows = diagValues.slice(1);

  pr30RequireHeaders_(summaryHeaders, ['신고연도','반기','사업자등록번호','구매카드사','카드매칭상태','주문건수','순수매출액','매입금액','납부예상부가세']);
  pr30RequireHeaders_(diagHeaders, ['신고연도','반기','주문번호','카드매칭상태','카드매칭근거','v6.69 2차귀속']);

  const diagStats = pr30DiagnosticStats_(diagHeaders, diagRows);
  const expectedDiag = {orders:1355, matched:766, nonCard:494, ambiguous:24, noMatch:71, fallback:593};
  Object.keys(expectedDiag).forEach(function(key) {
    if (diagStats[key] !== expectedDiag[key]) throw new Error('PR29 검증표 상태 불일치: ' + key + ' 실제 ' + diagStats[key] + ' / 기대 ' + expectedDiag[key]);
  });
  if (diagStats.invalidIdentity || diagStats.invalidFallbackEvidence) {
    throw new Error('PR29 검증표 안전검증 실패: 카드식별자 ' + diagStats.invalidIdentity + ' / 2차귀속 증빙 ' + diagStats.invalidFallbackEvidence);
  }

  const totals = pr30SummaryTotals_(summaryHeaders, summaryRows);
  const expectedTotals = {
    orders:1355, sales:71838700, salesSupply:65307938, salesVat:6530762,
    settlement:64726771, fee:7111929, purchase:54807644,
    purchaseSupply:49825146, purchaseVat:4982498, payable:1548264,
    profit:9919127, vatProfit:8370863
  };
  Object.keys(expectedTotals).forEach(function(key) {
    if (Math.round(totals[key]) !== expectedTotals[key]) throw new Error('PR29 요약 합계 불일치: ' + key + ' 실제 ' + totals[key] + ' / 기대 ' + expectedTotals[key]);
  });

  return {
    status:status,
    summaryHeaders:summaryHeaders,
    summaryRows:summaryRows,
    diagValues:diagValues,
    orders:diagStats.orders,
    matched:diagStats.matched,
    nonCard:diagStats.nonCard,
    ambiguous:diagStats.ambiguous,
    noMatch:diagStats.noMatch,
    fallback:diagStats.fallback,
    totals:totals
  };
}

function pr30ReadProduction_(ss) {
  const periodSheet = ss.getSheetByName(PR30_PRODUCTION_PERIOD_SHEET);
  const diagSheet = ss.getSheetByName(PR30_PRODUCTION_DIAG_SHEET) || ss.insertSheet(PR30_PRODUCTION_DIAG_SHEET);
  if (!periodSheet || periodSheet.getLastRow() < 2) throw new Error(PR30_PRODUCTION_PERIOD_SHEET + ' 시트가 없습니다.');

  const periodValues = periodSheet.getDataRange().getValues();
  const detailHeaderIndex = periodValues.findIndex(function(row, index) {
    return index >= 2 && pr30Text_(row[0]) === '집계구분' && row.map(pr30Text_).indexOf('신고연도') >= 0;
  });
  if (detailHeaderIndex < 0) throw new Error('부가세_기간별의 기간 상세표 시작행을 찾지 못했습니다.');

  const periodDetailValues = periodValues.slice(detailHeaderIndex);
  const diagValues = diagSheet.getLastRow() ? diagSheet.getDataRange().getValues() : [['']];
  return { periodSheet:periodSheet, diagSheet:diagSheet, periodValues:periodValues, diagValues:diagValues, periodDetailValues:periodDetailValues };
}

function pr30CreateBackupSheets_(ss, periodSheet, diagSheet) {
  [PR30_BACKUP_PERIOD_SHEET, PR30_BACKUP_DIAG_SHEET].forEach(function(name) {
    const old = ss.getSheetByName(name);
    if (old) ss.deleteSheet(old);
  });
  periodSheet.copyTo(ss).setName(PR30_BACKUP_PERIOD_SHEET);
  diagSheet.copyTo(ss).setName(PR30_BACKUP_DIAG_SHEET);
}

function pr30ApplyPeriodSummary_(sheet, headers, rows, detailValues) {
  const detailStartRow = rows.length + 5;
  const maxCols = Math.max(headers.length, detailValues[0] ? detailValues[0].length : 1, 1);
  sheet.clearContents();
  sheet.getRange(1, 1).setValue('사업자별 반기 신고요약 (구매카드별)');
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sheet.getRange(3, 1, rows.length, headers.length).setValues(rows);
  if (detailValues.length) sheet.getRange(detailStartRow, 1, detailValues.length, detailValues[0].length).setValues(detailValues);

  sheet.getRange(1, 1, 1, maxCols).setBackground('#b4c6e7').setFontWeight('bold');
  sheet.getRange(2, 1, 1, headers.length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
  if (detailValues.length) sheet.getRange(detailStartRow, 1, 1, detailValues[0].length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');

  const moneyHeaders = ['주문건수','순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익'];
  headers.forEach(function(header, index) {
    if (rows.length && moneyHeaders.indexOf(pr30Text_(header)) >= 0) sheet.getRange(3, index + 1, rows.length, 1).setNumberFormat('#,##0');
    else if (rows.length) sheet.getRange(3, index + 1, rows.length, 1).setNumberFormat('@');
    sheet.setColumnWidth(index + 1, /근거|비고/.test(header) ? 180 : (/카드|사업자|계정/.test(header) ? 135 : (/금액|부가세|이익|수수료/.test(header) ? 110 : 90)));
  });
  sheet.setFrozenRows(2);
}

function pr30ApplyDiagnostic_(sheet, values) {
  sheet.clearContents();
  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, values[0].length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
  const headers = values[0].map(pr30Text_);
  const purchaseIx = headers.indexOf('주문매입금액');
  const approvalIx = headers.indexOf('승인금액');
  if (values.length > 1 && purchaseIx >= 0) sheet.getRange(2, purchaseIx + 1, values.length - 1, 1).setNumberFormat('#,##0');
  if (values.length > 1 && approvalIx >= 0) sheet.getRange(2, approvalIx + 1, values.length - 1, 1).setNumberFormat('#,##0');
  headers.forEach(function(header, index) {
    sheet.setColumnWidth(index + 1, /후보요약|취소|근거|원본파일/.test(header) ? 220 : (/카드|사업자|계정|주문번호/.test(header) ? 135 : (/금액/.test(header) ? 105 : 90)));
  });
}

function pr30ValidateProductionAfterWrite_(ss, preview, originalDetailValues) {
  const periodSheet = ss.getSheetByName(PR30_PRODUCTION_PERIOD_SHEET);
  const diagSheet = ss.getSheetByName(PR30_PRODUCTION_DIAG_SHEET);
  const periodValues = periodSheet.getDataRange().getValues();
  const diagValues = diagSheet.getDataRange().getValues();

  const detailHeaderIndex = periodValues.findIndex(function(row, index) {
    return index >= 2 && pr30Text_(row[0]) === '집계구분' && row.map(pr30Text_).indexOf('신고연도') >= 0;
  });
  if (detailHeaderIndex < 0) throw new Error('적용 후 기간 상세표를 찾지 못했습니다.');
  const afterDetail = periodValues.slice(detailHeaderIndex);
  if (pr30MatrixSignature_(afterDetail) !== pr30MatrixSignature_(originalDetailValues)) throw new Error('기간 상세표가 변경되었습니다.');

  const summaryHeaders = periodValues[1].map(pr30Text_);
  const summaryRows = periodValues.slice(2, 2 + preview.summaryRows.length);
  if (pr30MatrixSignature_([summaryHeaders].concat(summaryRows)) !== pr30MatrixSignature_([preview.summaryHeaders].concat(preview.summaryRows))) {
    throw new Error('운영 구매카드 요약이 PR29 미리보기와 일치하지 않습니다.');
  }
  if (pr30MatrixSignature_(diagValues) !== pr30MatrixSignature_(preview.diagValues)) throw new Error('운영 카드매칭검증이 PR29 미리보기와 일치하지 않습니다.');

  const stats = pr30DiagnosticStats_(diagValues[0].map(pr30Text_), diagValues.slice(1));
  return {
    version:PR30_APPLY_VERSION,
    status:'PASS',
    orders:stats.orders,
    matched:stats.matched,
    nonCard:stats.nonCard,
    ambiguous:stats.ambiguous,
    noMatch:stats.noMatch,
    fallback:stats.fallback,
    sales:preview.totals.sales,
    purchase:preview.totals.purchase,
    payable:preview.totals.payable,
    detailRows:afterDetail.length - 1,
    completedAt:new Date().toISOString()
  };
}

function pr30DiagnosticStats_(headers, rows) {
  const statusIx = headers.indexOf('카드매칭상태');
  const fallbackIx = headers.indexOf('v6.69 2차귀속');
  const companyIx = headers.indexOf('구매카드사');
  const nameIx = headers.indexOf('구매카드명');
  const numberIx = headers.indexOf('카드번호');
  const end4Ix = headers.indexOf('카드번호끝4');
  const approvalDateIx = headers.indexOf('승인일');
  const approvalNoIx = headers.indexOf('승인번호');
  const approvalAmountIx = headers.indexOf('승인금액');
  const reasonIx = headers.indexOf('카드매칭근거');
  if ([statusIx,fallbackIx,companyIx,nameIx,numberIx,end4Ix,approvalDateIx,approvalNoIx,approvalAmountIx,reasonIx].some(function(i){return i < 0;})) {
    throw new Error('PR29 카드매칭검증 필수 헤더가 누락되었습니다.');
  }

  const stats = {orders:rows.length,matched:0,nonCard:0,ambiguous:0,noMatch:0,fallback:0,invalidIdentity:0,invalidFallbackEvidence:0};
  rows.forEach(function(row) {
    const status = pr30Text_(row[statusIx]);
    if (status === 'MATCHED' || status === 'MASTER_MATCHED') stats.matched++;
    else if (status === 'NON_CARD') stats.nonCard++;
    else if (status === 'AMBIGUOUS') stats.ambiguous++;
    else stats.noMatch++;

    const fallback = pr30Text_(row[fallbackIx]) === 'Y';
    if (fallback) {
      stats.fallback++;
      const invalidEvidence = pr30Text_(row[approvalDateIx]) || pr30Text_(row[approvalNoIx]) || pr30Number_(row[approvalAmountIx]) !== 0 ||
        (status !== 'AMBIGUOUS' && pr30Text_(row[reasonIx]).indexOf('금액비교없음') < 0);
      if (invalidEvidence) stats.invalidFallbackEvidence++;
    }

    const company = pr30Text_(row[companyIx]);
    const name = pr30Text_(row[nameIx]);
    const end4 = pr30End4_(row[end4Ix], row[numberIx]);
    if ((/KB|국민/.test(company) && end4 !== '4091') ||
        (/우리/.test(company) && end4 !== '7680') ||
        (name === 'Trip to 로카' && end4 !== '0126') ||
        (name === 'LOCA LIKIT 1.2' && end4 !== '0036')) stats.invalidIdentity++;
  });
  return stats;
}

function pr30SummaryTotals_(headers, rows) {
  const names = {
    orders:'주문건수', sales:'순수매출액', salesSupply:'매출공급가액', salesVat:'매출부가세',
    settlement:'정산기준금액', fee:'마켓수수료', purchase:'매입금액', purchaseSupply:'매입공급가액',
    purchaseVat:'매입부가세', payable:'납부예상부가세', profit:'예상이익', vatProfit:'부가세반영예상이익'
  };
  const indexes = {};
  Object.keys(names).forEach(function(key){ indexes[key] = headers.indexOf(names[key]); if (indexes[key] < 0) throw new Error('PR29 요약 필수 헤더 누락: ' + names[key]); });
  const totals = {};
  Object.keys(indexes).forEach(function(key){ totals[key] = 0; });
  rows.forEach(function(row){ Object.keys(indexes).forEach(function(key){ totals[key] += pr30Number_(row[indexes[key]]); }); });
  return totals;
}

function pr30WriteApplyStatus_(ss, status, message, data) {
  const sheet = ss.getSheetByName(PR30_STATUS_SHEET) || ss.insertSheet(PR30_STATUS_SHEET);
  const rows = [
    ['항목','값'],['버전',PR30_APPLY_VERSION],['상태',status],['메시지',message],
    ['운영 반영 대상','부가세_기간별 / 부가세_카드매칭검증'],
    ['백업 시트',PR30_BACKUP_PERIOD_SHEET + ' / ' + PR30_BACKUP_DIAG_SHEET]
  ];
  if (data) {
    [['상반기 주문','orders'],['MATCHED','matched'],['NON_CARD','nonCard'],['AMBIGUOUS','ambiguous'],['NO_MATCH','noMatch'],['2차귀속','fallback'],['순수매출액','sales'],['매입금액','purchase'],['납부예상부가세','payable'],['완료시각','completedAt']].forEach(function(pair){
      if (data[pair[1]] !== undefined) rows.push([pair[0], data[pair[1]]]);
      else if (data.totals && data.totals[pair[1]] !== undefined) rows.push([pair[0], data.totals[pair[1]]]);
    });
  }
  sheet.clearContents(); sheet.getRange(1,1,rows.length,2).setValues(rows); sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold'); sheet.setColumnWidth(1,220); sheet.setColumnWidth(2,520);
}

function pr30RestoreValues_(ss, name, values) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clearContents();
  sheet.getRange(1,1,values.length,values[0].length).setValues(values);
}

function pr30KeyValueMap_(values) {
  const out = {};
  (values || []).forEach(function(row){ const key = pr30Text_(row[0]); if (key) out[key] = row[1]; });
  return out;
}

function pr30RequireHeaders_(headers, required) {
  required.forEach(function(header){ if (headers.indexOf(header) < 0) throw new Error('필수 헤더 누락: ' + header); });
}

function pr30MatrixSignature_(values) {
  return JSON.stringify((values || []).map(function(row){ return (row || []).map(function(value){
    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value.toISOString();
    return value == null ? '' : value;
  }); }));
}

function pr30End4_(end4, number) {
  const explicit = String(end4 == null ? '' : end4).replace(/\D/g, '');
  if (explicit) return ('0000' + explicit).slice(-4);
  const digits = String(number == null ? '' : number).replace(/\D/g, '');
  return digits.length >= 4 ? digits.slice(-4) : '';
}

function pr30Text_(value) { return String(value == null ? '' : value).trim(); }
function pr30Number_(value) {
  if (typeof value === 'number') return isNaN(value) ? 0 : value;
  const n = Number(String(value == null ? '0' : value).replace(/[원,\s]/g, ''));
  return isNaN(n) ? 0 : n;
}
