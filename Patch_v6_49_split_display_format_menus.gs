/**
 * LOTTEON v6.49 split display format menus
 *
 * 목적:
 * - 표시 형식 정리를 메뉴별로 분리합니다.
 * - 금액/금액성 컬럼: 1,000단위 쉼표, 통화기호 없음 (#,##0)
 * - 율/비율 컬럼: 백분율 %, 소수점 첫째자리 (0.0%)
 *
 * 메뉴 함수:
 * - applyAmountThousandsFormat_v649
 * - applyPercentOneDecimalFormat_v649
 */

var LOTTEON_PATCH_V649_SPLIT_DISPLAY_FORMAT_MENUS_LOADED = true;

function applyAmountThousandsFormat_v649() {
  var started = Date.now();
  var result = applySplitDisplayFormatAllSheets_v649_('amount');
  try {
    SpreadsheetApp.getUi().alert(
      '금액 1,000단위 쉼표 적용 완료\n\n' +
      '처리 시트: ' + result.sheets + '\n' +
      '처리 컬럼: ' + result.columns + '\n' +
      '소요초: ' + Math.round((Date.now() - started) / 1000) + '\n\n' +
      '서식: #,##0\n통화기호 없음'
    );
  } catch (e) {}
  return result;
}

function applyPercentOneDecimalFormat_v649() {
  var started = Date.now();
  var result = applySplitDisplayFormatAllSheets_v649_('percent');
  try {
    SpreadsheetApp.getUi().alert(
      '백분율 % 표시 적용 완료\n\n' +
      '처리 시트: ' + result.sheets + '\n' +
      '처리 컬럼/행: ' + result.columns + '\n' +
      '소요초: ' + Math.round((Date.now() - started) / 1000) + '\n\n' +
      '서식: 0.0%'
    );
  } catch (e) {}
  return result;
}

function applySplitDisplayFormatAllSheets_v649_(mode) {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  var totalSheets = 0;
  var totalColumns = 0;
  sheets.forEach(function(sheet) {
    var result = applySplitDisplayFormatToSheet_v649_(sheet, mode);
    if (result.columns > 0) {
      totalSheets++;
      totalColumns += result.columns;
    }
  });
  try { log_('patch_v649_' + mode + '_format', 'sheets=' + totalSheets + ' columns=' + totalColumns); } catch (e) {}
  return { ok: true, mode: mode, sheets: totalSheets, columns: totalColumns };
}

function applySplitDisplayFormatToSheet_v649_(sheet, mode) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return { columns: 0 };
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headerScanRows = Math.min(lastRow, 120);
  var headersArea = sheet.getRange(1, 1, headerScanRows, lastCol).getValues();
  var processed = 0;
  var processedKeys = {};

  // 일반 표 및 대시보드 중간 섹션: 헤더 행을 찾아 해당 섹션만 적용합니다.
  for (var r = 1; r <= headerScanRows; r++) {
    var row = headersArea[r - 1];
    if (!isLikelyHeaderRow_v649_(row, mode)) continue;
    var endRow = findNextHeaderLikeRow_v649_(headersArea, r + 1, headerScanRows, mode) || (lastRow + 1);
    var rowCount = Math.max(0, Math.min(endRow, lastRow + 1) - r - 1);
    if (rowCount <= 0) continue;

    for (var c = 1; c <= lastCol; c++) {
      var header = normalizeHeader_v649_(row[c - 1]);
      var should = mode === 'amount' ? isAmountHeader_v649_(header) : isPercentHeader_v649_(header);
      if (!should) continue;
      var key = r + ':' + c + ':' + rowCount;
      if (processedKeys[key]) continue;
      processedKeys[key] = true;
      applyFormatRange_v649_(sheet.getRange(r + 1, c, rowCount, 1), mode);
      processed++;
    }
  }

  // 대시보드 요약형: B열 항목명, C열 값.
  if (lastCol >= 3) {
    var summaryRows = [];
    for (var sr = 1; sr <= Math.min(lastRow, 80); sr++) {
      var item = normalizeHeader_v649_(headersArea[sr - 1] && headersArea[sr - 1][1]);
      var shouldSummary = mode === 'amount' ? isAmountSummaryItem_v649_(item) : isPercentHeader_v649_(item);
      if (shouldSummary) summaryRows.push(sr);
    }
    summaryRows.forEach(function(rowNo) {
      applyFormatRange_v649_(sheet.getRange(rowNo, 3, 1, 1), mode);
      processed++;
    });
  }

  return { columns: processed };
}

function applyFormatRange_v649_(range, mode) {
  if (mode === 'amount') {
    range.setNumberFormat('#,##0').setHorizontalAlignment('right');
    return;
  }
  normalizePercentRangeValues_v649_(range);
  range.setNumberFormat('0.0%').setHorizontalAlignment('right');
}

function normalizePercentRangeValues_v649_(range) {
  var values = range.getValues();
  var changed = false;
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var before = values[r][c];
      var after = normalizePercentValue_v649_(before);
      if (after !== before) {
        values[r][c] = after;
        changed = true;
      }
    }
  }
  if (changed) range.setValues(values);
}

function normalizePercentValue_v649_(value) {
  if (value === '' || value == null) return value;
  if (typeof value === 'number') {
    if (!isFinite(value)) return value;
    if (Math.abs(value) > 1 && Math.abs(value) <= 100) return value / 100;
    return value;
  }
  var text = String(value).trim();
  if (!text) return value;
  var hadPercent = text.indexOf('%') >= 0;
  var cleaned = text.replace(/%/g, '').replace(/,/g, '').trim();
  if (!/^-?\d+(\.\d+)?$/.test(cleaned)) return value;
  var n = Number(cleaned);
  if (!isFinite(n)) return value;
  if (hadPercent) return n / 100;
  if (Math.abs(n) > 1 && Math.abs(n) <= 100) return n / 100;
  return n;
}

function isLikelyHeaderRow_v649_(row, mode) {
  var hit = 0;
  for (var i = 0; i < row.length; i++) {
    var h = normalizeHeader_v649_(row[i]);
    if (!h) continue;
    if (mode === 'amount' && isAmountHeader_v649_(h)) hit++;
    if (mode === 'percent' && isPercentHeader_v649_(h)) hit++;
  }
  return hit > 0;
}

function findNextHeaderLikeRow_v649_(rows, startRowNo, maxRowNo, mode) {
  for (var r = startRowNo; r <= maxRowNo; r++) {
    if (isLikelyHeaderRow_v649_(rows[r - 1], mode)) return r;
  }
  return 0;
}

function normalizeHeader_v649_(v) {
  return String(v == null ? '' : v).replace(/\n/g, '').replace(/\s+/g, '').trim();
}

function isAmountHeader_v649_(header) {
  header = normalizeHeader_v649_(header);
  if (!header) return false;
  if (/율|비율|rate|%/i.test(header)) return false;
  if (/상품수|건수|주문수|고객수|수량|행수|일수|경과|날짜|일자|번호|ID|아이디|URL|링크/.test(header)) return false;
  return /금액|매출|정산|매입|이익|수수료|부가세|공급가액|공급대가|구매가격|판매금액|결제금액|순수매출|총매출|예상정산|납부예상/.test(header);
}

function isAmountSummaryItem_v649_(item) {
  item = normalizeHeader_v649_(item);
  if (!item) return false;
  if (/율|비율|rate|%/.test(item)) return false;
  return /매출|정산|매입|이익|수수료|부가세|공급가액|공급대가|구매가격|금액/.test(item);
}

function isPercentHeader_v649_(header) {
  header = normalizeHeader_v649_(header);
  if (!header) return false;
  if (/금액|매출액|정산금액|매입금액|이익금액|수수료금액|부가세|공급가액|공급대가|수량|건수|상품수|주문수|고객수/.test(header)) return false;
  return /율|비율|rate|%/i.test(header);
}
