/**
 * PR38 operating hotfix v1.1
 *
 * - Copy the already validated temporary summary matrix directly into the
 *   production period sheet instead of reconstructing it with a second
 *   setValues pass.
 * - When matrix verification fails, report the first differing cell, value,
 *   and runtime type so the next correction does not require manual guessing.
 */
var PR38_HOTFIX_VERSION = 'v1.1-SUMMARY-COPY-DIFF';

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

  const sourceRange = summary.getDataRange();
  const summaryRowsIncludingHeader = sourceRange.getNumRows();
  const summaryCols = sourceRange.getNumColumns();
  if (summaryRowsIncludingHeader !== state.summaryRows + 1) {
    throw new Error(
      '요약 준비 행 수 불일치: 실제 ' + summaryRowsIncludingHeader +
      ' / 기대 ' + (state.summaryRows + 1)
    );
  }

  period.clear();
  period.getRange(1, 1).setValue('사업자별 반기 신고요약 (구매카드별)');
  sourceRange.copyTo(
    period.getRange(2, 1, summaryRowsIncludingHeader, summaryCols),
    SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
    false
  );

  period.getRange(1, 1, 1, Math.max(summaryCols, state.periodDetailCols))
    .setBackground('#b4c6e7').setFontWeight('bold');
  period.getRange(2, 1, 1, summaryCols)
    .setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');

  if (state.summaryRows) {
    period.getRange(3, 1, state.summaryRows, 11).setNumberFormat('@');
    for (let c = 12; c <= 23; c++) {
      period.getRange(3, c, state.summaryRows, 1).setNumberFormat('#,##0');
    }
  }
  period.setFrozenRows(2);
  return pr38Next_(
    ss,
    state,
    'WRITE_PERIOD',
    '부가세_기간별 요약 원본복사 준비 완료 (' + PR38_HOTFIX_VERSION + ')'
  );
}

function pr38AssertMatrix_(label, expected, actual) {
  const expectedRows = expected || [];
  const actualRows = actual || [];
  if (expectedRows.length !== actualRows.length) {
    throw new Error(
      label + ' 행 수 불일치: 실제 ' + actualRows.length + ' / 기대 ' + expectedRows.length
    );
  }

  for (let r = 0; r < expectedRows.length; r++) {
    const expectedRow = expectedRows[r] || [];
    const actualRow = actualRows[r] || [];
    if (expectedRow.length !== actualRow.length) {
      throw new Error(
        label + ' 열 수 불일치: 행 ' + (r + 1) +
        ' 실제 ' + actualRow.length + ' / 기대 ' + expectedRow.length
      );
    }
    for (let c = 0; c < expectedRow.length; c++) {
      const left = pr38CanonicalCellHotfix_(expectedRow[c]);
      const right = pr38CanonicalCellHotfix_(actualRow[c]);
      if (left !== right) {
        throw new Error(
          label + ' 불일치 R' + (r + 1) + 'C' + (c + 1) +
          ' 기대[' + pr38DescribeCellHotfix_(expectedRow[c]) + ']' +
          ' 실제[' + pr38DescribeCellHotfix_(actualRow[c]) + ']'
        );
      }
    }
  }
}

function pr38CanonicalCellHotfix_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return 'DATE:' + value.toISOString();
  }
  if (typeof value === 'number') return 'NUM:' + String(value);
  if (typeof value === 'boolean') return 'BOOL:' + String(value);
  return 'TEXT:' + pr38Text_(value);
}

function pr38DescribeCellHotfix_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return 'Date:' + value.toISOString();
  }
  return (typeof value) + ':' + pr38Text_(value);
}
