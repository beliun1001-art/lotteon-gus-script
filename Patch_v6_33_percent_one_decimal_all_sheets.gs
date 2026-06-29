/**
 * LOTTEON v6.33 percent one-decimal format for all sheets
 *
 * 확정 기준:
 * - 모든 시트에서 율/비율/rate/% 컬럼은 퍼센트 형식으로 표시합니다.
 * - 소수점은 첫째 자리까지 표시합니다. 예: 17.1%
 * - 값이 17.1 또는 "17.1%"처럼 들어온 경우 0.171로 보정한 뒤 17.1%로 표시합니다.
 * - 값이 0.171처럼 들어온 경우 값은 유지하고 17.1%로 표시합니다.
 *
 * 적용 경로:
 * - 대시보드 빠른 갱신 후
 * - 부가세 신고자료 생성 후
 * - 표시서식만 빠른 정리 실행 시
 * - 열너비 자동조정 실행 후
 */

var LOTTEON_PATCH_V633_PERCENT_ONE_DECIMAL_ALL_SHEETS_LOADED = true;

var __baseApplyDisplayStandardsOnlyFast_v633 = typeof applyDisplayStandardsOnlyFast_v623 === 'function' ? applyDisplayStandardsOnlyFast_v623 : null;
var __baseApplyFastOutputSheetFormatting_v633 = typeof applyFastOutputSheetFormatting_v620_ === 'function' ? applyFastOutputSheetFormatting_v620_ : null;
var __baseFormatOutputSheetBasic_v633 = typeof formatOutputSheetBasic_v628_ === 'function' ? formatOutputSheetBasic_v628_ : null;
var __baseHardFixDashboardDisplay_v633 = typeof hardFixDashboardDisplay_v630_ === 'function' ? hardFixDashboardDisplay_v630_ : null;
var __baseForceDashboardFinalAlignmentWidth_v633 = typeof forceDashboardFinalAlignmentWidth_v631_ === 'function' ? forceDashboardFinalAlignmentWidth_v631_ : null;
var __baseRunColumnWidthStep_v633 = typeof runColumnWidthAutoAdjustStep_v623 === 'function' ? runColumnWidthAutoAdjustStep_v623 : null;
var __baseRefreshDashboardFastOnly_v633 = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;
var __baseGenerateVatReportsFullSeparated_v633 = typeof generateVatReportsFullSeparated_v622 === 'function' ? generateVatReportsFullSeparated_v622 : null;

applyDisplayStandardsOnlyFast_v623 = function() {
  var result = __baseApplyDisplayStandardsOnlyFast_v633 ? __baseApplyDisplayStandardsOnlyFast_v633.apply(this, arguments) : null;
  applyPercentOneDecimalAllSheets_v633_();
  try { SpreadsheetApp.getUi().alert('표시서식 정리 완료\n\n모든 시트의 율/비율 컬럼을 소수점 첫째자리 % 형식으로 정리했습니다.'); } catch (e) {}
  return result || { ok: true };
};

applyFastOutputSheetFormatting_v620_ = function() {
  var result = __baseApplyFastOutputSheetFormatting_v633 ? __baseApplyFastOutputSheetFormatting_v633.apply(this, arguments) : null;
  applyPercentOneDecimalAllSheets_v633_();
  return result || { ok: true };
};

formatOutputSheetBasic_v628_ = function(sheet) {
  var result = __baseFormatOutputSheetBasic_v633 ? __baseFormatOutputSheetBasic_v633.apply(this, arguments) : null;
  applyPercentOneDecimalToSheet_v633_(sheet);
  return result || { ok: true };
};

hardFixDashboardDisplay_v630_ = function() {
  var result = __baseHardFixDashboardDisplay_v633 ? __baseHardFixDashboardDisplay_v633.apply(this, arguments) : null;
  applyPercentOneDecimalToDashboard_v633_();
  return result || { ok: true };
};

forceDashboardFinalAlignmentWidth_v631_ = function() {
  var result = __baseForceDashboardFinalAlignmentWidth_v633 ? __baseForceDashboardFinalAlignmentWidth_v633.apply(this, arguments) : null;
  applyPercentOneDecimalToDashboard_v633_();
  return result || { ok: true };
};

runColumnWidthAutoAdjustStep_v623 = function() {
  var result = __baseRunColumnWidthStep_v633 ? __baseRunColumnWidthStep_v633.apply(this, arguments) : null;
  applyPercentOneDecimalAllSheets_v633_();
  return result;
};

refreshDashboardFastOnly = function() {
  var result = __baseRefreshDashboardFastOnly_v633 ? __baseRefreshDashboardFastOnly_v633.apply(this, arguments) : null;
  applyPercentOneDecimalAllSheets_v633_();
  return result;
};

generateVatReportsFullSeparated_v622 = function() {
  var result = __baseGenerateVatReportsFullSeparated_v633 ? __baseGenerateVatReportsFullSeparated_v633.apply(this, arguments) : null;
  applyPercentOneDecimalAllSheets_v633_();
  return result;
};

function applyPercentOneDecimalAllSheets_v633_() {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  var processed = 0;
  sheets.forEach(function(sheet) {
    if (applyPercentOneDecimalToSheet_v633_(sheet).ok) processed++;
  });
  try { log_('patch_v633_percent_one_decimal_all_sheets', 'processed=' + processed); } catch (e) {}
  return { ok: true, processed: processed };
}

function applyPercentOneDecimalToDashboard_v633_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!sheet) return { skipped: true };
  return applyPercentOneDecimalToSheet_v633_(sheet);
}

function applyPercentOneDecimalToSheet_v633_(sheet) {
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return { skipped: true };
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var maxRows = Math.min(lastRow, 5000);
  var values = sheet.getRange(1, 1, maxRows, lastCol).getValues();

  // 일반 표: 1행 헤더 기준
  applyPercentColumnsByHeaderRow_v633_(sheet, values, 1, maxRows, lastCol);

  // 대시보드/섹션형 표: 중간중간 '구분' 헤더 행 기준
  for (var r = 1; r <= maxRows; r++) {
    var first = String(values[r - 1][0] == null ? '' : values[r - 1][0]).replace(/\n/g, '').trim();
    if (r > 1 && first === '구분') {
      applyPercentColumnsByHeaderRow_v633_(sheet, values, r, maxRows, lastCol);
    }
  }

  // 대시보드 요약형: B열 항목명, C열 값
  applyPercentSummaryRows_v633_(sheet, values, maxRows);

  return { ok: true };
}

function applyPercentColumnsByHeaderRow_v633_(sheet, values, headerRowNo, maxRows, lastCol) {
  if (headerRowNo >= maxRows) return;
  var endRow = findNextSectionHeaderRow_v633_(values, headerRowNo + 1, maxRows) || (maxRows + 1);
  var rowCount = endRow - headerRowNo - 1;
  if (rowCount <= 0) return;

  for (var c = 1; c <= lastCol; c++) {
    var header = normalizeHeader_v633_(values[headerRowNo - 1][c - 1]);
    if (!isRateHeader_v633_(header)) continue;
    var range = sheet.getRange(headerRowNo + 1, c, rowCount, 1);
    normalizePercentRangeValues_v633_(range);
    range.setNumberFormat('0.0%').setHorizontalAlignment('right');
  }
}

function applyPercentSummaryRows_v633_(sheet, values, maxRows) {
  var limit = Math.min(maxRows, 80);
  for (var r = 1; r <= limit; r++) {
    var item = normalizeHeader_v633_(values[r - 1][1]);
    if (!isRateHeader_v633_(item)) continue;
    var cell = sheet.getRange(r, 3);
    normalizePercentRangeValues_v633_(cell);
    cell.setNumberFormat('0.0%').setHorizontalAlignment('right');
  }
}

function normalizePercentRangeValues_v633_(range) {
  var values = range.getValues();
  var changed = false;
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var before = values[r][c];
      var after = normalizePercentValue_v633_(before);
      if (after !== before) {
        values[r][c] = after;
        changed = true;
      }
    }
  }
  if (changed) range.setValues(values);
}

function normalizePercentValue_v633_(value) {
  if (value === '' || value == null) return value;
  if (typeof value === 'number') {
    if (!isFinite(value)) return value;
    // 1 초과 100 이하 숫자는 17.1 → 17.1% 의도값으로 보고 0.171로 보정합니다.
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

function findNextSectionHeaderRow_v633_(values, startRowNo, maxRows) {
  for (var r = startRowNo; r <= maxRows; r++) {
    var first = String(values[r - 1][0] == null ? '' : values[r - 1][0]).replace(/\n/g, '').trim();
    if (first === '구분') return r;
  }
  return 0;
}

function normalizeHeader_v633_(v) {
  return String(v == null ? '' : v).replace(/\n/g, '').replace(/\s+/g, '').trim();
}

function isRateHeader_v633_(header) {
  header = normalizeHeader_v633_(header);
  if (!header) return false;
  if (/금액|매출액|정산금액|매입금액|이익금액|수수료금액|부가세|공급가액|공급대가/.test(header)) return false;
  return /율|비율|rate|%/i.test(header);
}
