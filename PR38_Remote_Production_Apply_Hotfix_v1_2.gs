/**
 * PR38 operating hotfix v1.2
 *
 * The production summary intentionally formats columns A:K as text because
 * they are identifiers and labels. Google Sheets can therefore return the
 * same visible year as number 2026 in the temporary sheet and string "2026"
 * in the production sheet. Treat only those summary identifier columns as
 * display-text equivalents; keep all financial columns and all other matrix
 * verification type-strict.
 */
var PR38_HOTFIX_V12_VERSION = 'v1.2-SUMMARY-ID-TEXT-EQUIVALENCE';

function pr38AssertMatrix_(label, expected, actual) {
  const expectedRows = expected || [];
  const actualRows = actual || [];
  if (expectedRows.length !== actualRows.length) {
    throw new Error(
      label + ' 행 수 불일치: 실제 ' + actualRows.length + ' / 기대 ' + expectedRows.length
    );
  }

  const isSummary = String(label || '').indexOf('구매카드별 반기 요약') === 0;

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
      const identifierTextColumn = isSummary && c < 11;
      const left = identifierTextColumn
        ? 'IDTEXT:' + pr38Text_(expectedRow[c])
        : pr38CanonicalCellHotfixV12_(expectedRow[c]);
      const right = identifierTextColumn
        ? 'IDTEXT:' + pr38Text_(actualRow[c])
        : pr38CanonicalCellHotfixV12_(actualRow[c]);

      if (left !== right) {
        throw new Error(
          label + ' 불일치 R' + (r + 1) + 'C' + (c + 1) +
          ' 기대[' + pr38DescribeCellHotfixV12_(expectedRow[c]) + ']' +
          ' 실제[' + pr38DescribeCellHotfixV12_(actualRow[c]) + ']' +
          ' / 검증=' + PR38_HOTFIX_V12_VERSION
        );
      }
    }
  }
}

function pr38CanonicalCellHotfixV12_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return 'DATE:' + value.toISOString();
  }
  if (typeof value === 'number') return 'NUM:' + String(value);
  if (typeof value === 'boolean') return 'BOOL:' + String(value);
  return 'TEXT:' + pr38Text_(value);
}

function pr38DescribeCellHotfixV12_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return 'Date:' + value.toISOString();
  }
  return (typeof value) + ':' + pr38Text_(value);
}
