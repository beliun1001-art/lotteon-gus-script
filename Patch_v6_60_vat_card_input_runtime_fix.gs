/** v6.60 Issue #20 runtime input/date hardening for purchase-card reconciliation. */
var LOTTEON_PATCH_V660_VAT_CARD_INPUT_RUNTIME_FIX_LOADED = true;

// Google Sheets may return actual Date objects for pasted/typed dates.
normalizeDateText_v660_ = function(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return v.getFullYear() + '-' + pad2_v660_(v.getMonth() + 1) + '-' + pad2_v660_(v.getDate());
  }
  var s = text_v660_(v);
  if (!s) return '';
  var m = s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + pad2_v660_(m[2]) + '-' + pad2_v660_(m[3]);
};

// Protect identifiers from number/scientific-notation conversion before the user pastes data.
ensureSheetHeaders_v660_ = function(ss, name, headers) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
  }
  var textHeaders = /카드번호|카드번호끝4|사업자코드|사업자등록번호|승인번호|가맹점주문번호/;
  var rows = Math.max((sheet.getMaxRows ? sheet.getMaxRows() : 1000) - 1, 1);
  headers.forEach(function(header, index) {
    if (textHeaders.test(String(header || ''))) sheet.getRange(2, index + 1, rows, 1).setNumberFormat('@');
  });
  return sheet;
};

// Public helpers for operating smoke / later routine maintenance.
function prepareVatCardInputs_v660_() {
  var ss = SpreadsheetApp.getActive();
  ensureVatCardInputSheets_v660_(ss);
  try { ss.toast('카드_마스터 / 카드사용내역_붙여넣기 입력 시트를 준비했습니다.', 'LOTTEON', 5); } catch (e) {}
  return { ok:true, masterSheet:LOTTEON_V660_CARD_MASTER_SHEET, historySheet:LOTTEON_V660_CARD_HISTORY_SHEET };
}

function rebuildVatPurchaseCardSummary_v660_() {
  var ss = SpreadsheetApp.getActive();
  ensureVatCardInputSheets_v660_(ss);
  var result = buildVatPurchaseCardReconciliation_v660_(ss);
  try { ss.toast('구매카드 매칭 및 부가세 신고요약을 갱신했습니다.', 'LOTTEON', 5); } catch (e) {}
  return result;
}
