/**
 * PR #31 operating fix v1.1.
 * Run only runPr30ApplyPreviewToProductionV11().
 * Requires PR30_Production_Apply.gs to remain installed beside this file.
 */
const PR30_V11_VERSION = 'v1.1-PR30-TEXT-FIRST-CANONICAL-COMPARE';
const PR30_V11_SUMMARY_NUMERIC = [
  '주문건수','순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료',
  '매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익'
];
const PR30_V11_DIAG_NUMERIC = ['상세행수','주문매입금액','승인금액','후보수'];
const PR30_V11_DETAIL_NUMERIC = [
  '주문건수','순수매출액','정산기준금액','마켓수수료','매입금액','매출부가세',
  '매입부가세','납부예상부가세','예상이익','부가세반영예상이익'
];

function runPr30ApplyPreviewToProductionV11() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  if (typeof pr30ReadAndValidatePreview_ !== 'function' || typeof pr30ReadProduction_ !== 'function') {
    throw new Error('PR30_Production_Apply.gs 원본 파일을 찾지 못했습니다.');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('다른 작업이 실행 중입니다. 잠시 후 다시 실행하세요.');

  let periodBefore = null;
  let diagBefore = null;
  try {
    pr30v11RecoverPriorRollback_(ss);
    const preview = pr30ReadAndValidatePreview_(ss);
    const production = pr30ReadProduction_(ss);
    periodBefore = production.periodValues;
    diagBefore = production.diagValues;

    pr30CreateBackupSheets_(ss, production.periodSheet, production.diagSheet);
    pr30v11WriteStatus_(ss, 'WRITING', 'v1.1 백업 완료; 운영 시트 반영 중', preview);

    pr30v11ApplySummary_(production.periodSheet, preview.summaryHeaders, preview.summaryRows, production.periodDetailValues);
    pr30v11ApplyDiagnostic_(production.diagSheet, preview.diagValues);
    SpreadsheetApp.flush();

    const verified = pr30v11ValidateAfterWrite_(ss, preview, production.periodDetailValues);
    pr30v11WriteStatus_(ss, 'PASS', 'v1.1 운영 반영 및 사후 검증 완료', verified);
    SpreadsheetApp.getUi().alert(
      'PR #31 운영 반영 완료\n\n' +
      '상반기 주문: ' + verified.orders + '건\n' +
      'MATCHED: ' + verified.matched + '건\n' +
      'NON_CARD: ' + verified.nonCard + '건\n' +
      'AMBIGUOUS: ' + verified.ambiguous + '건\n' +
      'NO_MATCH: ' + verified.noMatch + '건\n' +
      '2차귀속: ' + verified.fallback + '건'
    );
    return verified;
  } catch (error) {
    try {
      const periodBackup = ss.getSheetByName(PR30_BACKUP_PERIOD_SHEET);
      const diagBackup = ss.getSheetByName(PR30_BACKUP_DIAG_SHEET);
      if (periodBackup && diagBackup) {
        pr30v11RestoreBackup_(ss, PR30_BACKUP_PERIOD_SHEET, PR30_PRODUCTION_PERIOD_SHEET);
        pr30v11RestoreBackup_(ss, PR30_BACKUP_DIAG_SHEET, PR30_PRODUCTION_DIAG_SHEET);
      } else {
        if (periodBefore && periodBefore.length) pr30RestoreValues_(ss, PR30_PRODUCTION_PERIOD_SHEET, periodBefore);
        if (diagBefore && diagBefore.length) pr30RestoreValues_(ss, PR30_PRODUCTION_DIAG_SHEET, diagBefore);
      }
      SpreadsheetApp.flush();
      pr30v11WriteStatus_(ss, 'ROLLED_BACK', String(error && error.message ? error.message : error), null);
    } catch (rollbackError) {
      pr30v11WriteStatus_(ss, 'ROLLBACK_ERROR', String(rollbackError && rollbackError.message ? rollbackError.message : rollbackError), null);
    }
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function pr30v11RecoverPriorRollback_(ss) {
  const statusSheet = ss.getSheetByName(PR30_STATUS_SHEET);
  if (!statusSheet || statusSheet.getLastRow() < 2) return;
  const status = pr30KeyValueMap_(statusSheet.getDataRange().getValues());
  if (pr30Text_(status['상태']) !== 'ROLLED_BACK') return;
  if (!ss.getSheetByName(PR30_BACKUP_PERIOD_SHEET) || !ss.getSheetByName(PR30_BACKUP_DIAG_SHEET)) return;
  pr30v11RestoreBackup_(ss, PR30_BACKUP_PERIOD_SHEET, PR30_PRODUCTION_PERIOD_SHEET);
  pr30v11RestoreBackup_(ss, PR30_BACKUP_DIAG_SHEET, PR30_PRODUCTION_DIAG_SHEET);
  SpreadsheetApp.flush();
}

function pr30v11ApplySummary_(sheet, headers, rows, detailValues) {
  const h = headers.map(pr30Text_);
  const prepared = pr30v11Prepare_(h, rows, PR30_V11_SUMMARY_NUMERIC);
  const detailStart = prepared.length + 5;
  const maxCols = Math.max(h.length, detailValues[0] ? detailValues[0].length : 1, 1);

  sheet.clearContents();
  sheet.getRange(1,1).setValue('사업자별 반기 신고요약 (구매카드별)');
  sheet.getRange(2,1,1,h.length).setValues([h]);
  h.forEach(function(header, index) {
    if (prepared.length) sheet.getRange(3,index+1,prepared.length,1)
      .setNumberFormat(PR30_V11_SUMMARY_NUMERIC.indexOf(header) >= 0 ? '#,##0' : '@');
  });
  if (prepared.length) sheet.getRange(3,1,prepared.length,h.length).setValues(prepared);

  if (detailValues.length) {
    const dh = detailValues[0].map(pr30Text_);
    const dr = pr30v11Prepare_(dh, detailValues.slice(1), PR30_V11_DETAIL_NUMERIC);
    sheet.getRange(detailStart,1,1,dh.length).setValues([dh]);
    dh.forEach(function(header, index) {
      if (dr.length) sheet.getRange(detailStart+1,index+1,dr.length,1)
        .setNumberFormat(PR30_V11_DETAIL_NUMERIC.indexOf(header) >= 0 ? '#,##0' : '@');
    });
    if (dr.length) sheet.getRange(detailStart+1,1,dr.length,dh.length).setValues(dr);
  }

  sheet.getRange(1,1,1,maxCols).setBackground('#b4c6e7').setFontWeight('bold');
  sheet.getRange(2,1,1,h.length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
  if (detailValues.length) sheet.getRange(detailStart,1,1,detailValues[0].length)
    .setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setFrozenRows(2);
}

function pr30v11ApplyDiagnostic_(sheet, values) {
  const h = values[0].map(pr30Text_);
  const prepared = pr30v11Prepare_(h, values.slice(1), PR30_V11_DIAG_NUMERIC);
  sheet.clearContents();
  sheet.getRange(1,1,1,h.length).setValues([h]);
  h.forEach(function(header, index) {
    if (prepared.length) sheet.getRange(2,index+1,prepared.length,1)
      .setNumberFormat(PR30_V11_DIAG_NUMERIC.indexOf(header) >= 0 ? '#,##0' : '@');
  });
  if (prepared.length) sheet.getRange(2,1,prepared.length,h.length).setValues(prepared);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,h.length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
}

function pr30v11ValidateAfterWrite_(ss, preview, originalDetailValues) {
  const period = ss.getSheetByName(PR30_PRODUCTION_PERIOD_SHEET).getDataRange().getValues();
  const detailHeader = period.findIndex(function(row, index) {
    return index >= 2 && pr30Text_(row[0]) === '집계구분' && row.map(pr30Text_).indexOf('신고연도') >= 0;
  });
  if (detailHeader < 0) throw new Error('적용 후 기간 상세표를 찾지 못했습니다.');

  const detailWidth = originalDetailValues[0].length;
  const actualDetail = period.slice(detailHeader, detailHeader + originalDetailValues.length)
    .map(function(row) { return row.slice(0, detailWidth); });
  pr30v11AssertEqual_(
    '기간 상세표',
    pr30v11Canonical_(originalDetailValues[0].map(pr30Text_), originalDetailValues.slice(1), PR30_V11_DETAIL_NUMERIC),
    pr30v11Canonical_(actualDetail[0].map(pr30Text_), actualDetail.slice(1), PR30_V11_DETAIL_NUMERIC)
  );

  const summaryWidth = preview.summaryHeaders.length;
  const actualSummaryHeaders = period[1].slice(0,summaryWidth).map(pr30Text_);
  const actualSummaryRows = period.slice(2,2+preview.summaryRows.length)
    .map(function(row) { return row.slice(0,summaryWidth); });
  pr30v11AssertEqual_(
    '운영 구매카드 요약',
    pr30v11Canonical_(preview.summaryHeaders, preview.summaryRows, PR30_V11_SUMMARY_NUMERIC),
    pr30v11Canonical_(actualSummaryHeaders, actualSummaryRows, PR30_V11_SUMMARY_NUMERIC)
  );

  const diagHeaders = preview.diagValues[0].map(pr30Text_);
  const diagRows = preview.diagValues.slice(1);
  const actualDiag = ss.getSheetByName(PR30_PRODUCTION_DIAG_SHEET)
    .getRange(1,1,preview.diagValues.length,diagHeaders.length).getValues();
  pr30v11AssertEqual_(
    '운영 카드매칭검증',
    pr30v11Canonical_(diagHeaders, diagRows, PR30_V11_DIAG_NUMERIC),
    pr30v11Canonical_(actualDiag[0].map(pr30Text_), actualDiag.slice(1), PR30_V11_DIAG_NUMERIC)
  );

  const stats = pr30DiagnosticStats_(actualDiag[0].map(pr30Text_), actualDiag.slice(1));
  return {
    version:PR30_V11_VERSION,status:'PASS',orders:stats.orders,matched:stats.matched,
    nonCard:stats.nonCard,ambiguous:stats.ambiguous,noMatch:stats.noMatch,fallback:stats.fallback,
    sales:preview.totals.sales,purchase:preview.totals.purchase,payable:preview.totals.payable,
    completedAt:new Date().toISOString()
  };
}

function pr30v11Prepare_(headers, rows, numericHeaders) {
  const numeric = {};
  numericHeaders.forEach(function(header) { numeric[header] = true; });
  const numberIx = headers.indexOf('카드번호');
  return rows.map(function(row) {
    return headers.map(function(header,index) {
      if (numeric[header]) return pr30Number_(row[index]);
      if (header === '카드번호끝4') return pr30End4_(row[index], numberIx >= 0 ? row[numberIx] : '');
      return pr30Text_(row[index]);
    });
  });
}

function pr30v11Canonical_(headers, rows, numericHeaders) {
  const h = headers.map(pr30Text_);
  return [h].concat(pr30v11Prepare_(h, rows, numericHeaders));
}

function pr30v11AssertEqual_(label, expected, actual) {
  const maxRows = Math.max(expected.length, actual.length);
  for (let r=0; r<maxRows; r++) {
    const er = expected[r] || [], ar = actual[r] || [];
    const maxCols = Math.max(er.length, ar.length);
    for (let c=0; c<maxCols; c++) {
      const ev = er[c] === undefined ? '' : er[c];
      const av = ar[c] === undefined ? '' : ar[c];
      if (JSON.stringify(ev) !== JSON.stringify(av)) {
        throw new Error(label + ' 불일치: 행 ' + (r+1) + ' / 열 ' + (c+1) +
          ' / 기대 ' + JSON.stringify(ev) + ' / 실제 ' + JSON.stringify(av));
      }
    }
  }
}

function pr30v11RestoreBackup_(ss, backupName, targetName) {
  const backup = ss.getSheetByName(backupName);
  const target = ss.getSheetByName(targetName) || ss.insertSheet(targetName);
  if (!backup) throw new Error('백업 시트가 없습니다: ' + backupName);
  target.clear();
  backup.getDataRange().copyTo(target.getRange(1,1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  target.setFrozenRows(backup.getFrozenRows());
  target.setFrozenColumns(backup.getFrozenColumns());
  for (let c=1; c<=backup.getLastColumn(); c++) target.setColumnWidth(c, backup.getColumnWidth(c));
}

function pr30v11WriteStatus_(ss, status, message, data) {
  const sheet = ss.getSheetByName(PR30_STATUS_SHEET) || ss.insertSheet(PR30_STATUS_SHEET);
  const rows = [
    ['항목','값'],['버전',PR30_V11_VERSION],['상태',status],['메시지',message],
    ['운영 반영 대상','부가세_기간별 / 부가세_카드매칭검증'],
    ['백업 시트',PR30_BACKUP_PERIOD_SHEET + ' / ' + PR30_BACKUP_DIAG_SHEET]
  ];
  if (data) {
    [['상반기 주문','orders'],['MATCHED','matched'],['NON_CARD','nonCard'],['AMBIGUOUS','ambiguous'],
     ['NO_MATCH','noMatch'],['2차귀속','fallback'],['순수매출액','sales'],['매입금액','purchase'],
     ['납부예상부가세','payable'],['완료시각','completedAt']].forEach(function(pair) {
      if (data[pair[1]] !== undefined) rows.push([pair[0],data[pair[1]]]);
      else if (data.totals && data.totals[pair[1]] !== undefined) rows.push([pair[0],data.totals[pair[1]]]);
    });
  }
  sheet.clearContents();
  sheet.getRange(1,1,rows.length,2).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');
  sheet.setColumnWidth(1,220); sheet.setColumnWidth(2,560);
}
