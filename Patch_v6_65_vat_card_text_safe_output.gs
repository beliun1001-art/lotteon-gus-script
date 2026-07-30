/** v6.65 Issue #20: VAT purchase-card diagnostic/summary text-safe output only. */
var LOTTEON_PATCH_V665_VAT_CARD_TEXT_SAFE_OUTPUT_LOADED = true;

function normalizeVatCardDisplayTime_v665_(value) {
  var s = String(value == null ? '' : value).trim();
  if (!s) return '';
  var direct = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (direct) return (direct[1].length < 2 ? '0' + direct[1] : direct[1]) + ':' + direct[2] + ':' + (direct[3] || '00');
  var embedded = s.match(/\b(\d{2}:\d{2}:\d{2})\b/);
  return embedded ? embedded[1] : s;
}

function normalizeVatCardDisplayText_v665_(value) {
  var s = String(value == null ? '' : value);
  if (!s) return '';
  return s.replace(/\b(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+\w+\s+\d{1,2}\s+\d{4}\s+(\d{2}:\d{2}:\d{2})\s+GMT[+-]\d{4}\s+\([^)]+\)/g, '$1');
}

function prepareVatCardDiagnosticTextColumns_v665_(sheet, rowCount) {
  if (!sheet || !rowCount) return;
  // Text identifiers: business/account/order/card/approval/merchant-order plus approval time.
  [4,5,6,13,14,16,17,23].forEach(function(col) {
    sheet.getRange(2, col, rowCount, 1).setNumberFormat('@');
  });
}

function cleanVatCardDiagnosticDisplay_v665_(sheet, rowCount) {
  if (!sheet || !rowCount) return;
  var timeRange = sheet.getRange(2, 16, rowCount, 1);
  var timeValues = timeRange.getValues();
  for (var r = 0; r < timeValues.length; r++) timeValues[r][0] = normalizeVatCardDisplayTime_v665_(timeValues[r][0]);
  timeRange.setNumberFormat('@').setValues(timeValues);

  var summaryRange = sheet.getRange(2, 27, rowCount, 1);
  var summaryValues = summaryRange.getValues();
  for (var i = 0; i < summaryValues.length; i++) summaryValues[i][0] = normalizeVatCardDisplayText_v665_(summaryValues[i][0]);
  summaryRange.setValues(summaryValues);
}

var __baseWriteVatCardMatchDiagnostic_v665_ = typeof writeVatCardMatchDiagnostic_v660_ === 'function' ? writeVatCardMatchDiagnostic_v660_ : null;
if (__baseWriteVatCardMatchDiagnostic_v665_) {
  writeVatCardMatchDiagnostic_v660_ = function(ss, orders) {
    var rows = (orders || []).length;
    var sheet = ss && ss.getSheetByName && ss.getSheetByName(LOTTEON_V660_CARD_DIAG_SHEET);
    if (!sheet && ss && ss.insertSheet) sheet = ss.insertSheet(LOTTEON_V660_CARD_DIAG_SHEET);
    prepareVatCardDiagnosticTextColumns_v665_(sheet, rows);
    var result = __baseWriteVatCardMatchDiagnostic_v665_.apply(this, arguments);
    cleanVatCardDiagnosticDisplay_v665_(sheet, rows);
    return result;
  };
}

var __basePrependVatBusinessCardHalfSummary_v665_ = typeof prependVatBusinessCardHalfSummary_v660_ === 'function' ? prependVatBusinessCardHalfSummary_v660_ : null;
if (__basePrependVatBusinessCardHalfSummary_v665_) {
  prependVatBusinessCardHalfSummary_v660_ = function(sheet, periodValues, summary) {
    var rows = (summary || []).length;
    // Must be text BEFORE setValues so identifiers such as 036 / 7680 / 4091 are not numeric-formatted.
    if (sheet && rows) sheet.getRange(3, 1, rows, 11).setNumberFormat('@');
    return __basePrependVatBusinessCardHalfSummary_v665_.apply(this, arguments);
  };
}
