/** v6.73: hide completed cleanup rows while keeping their deletion log in the manager sheet. */
var LOTTEON_PATCH_V673_CLEANUP_HIDE_DELETED_ROWS_LOADED = true;

if (typeof buildLotteonSheetCleanupList_v671_ !== 'function' ||
    typeof deleteCheckedLotteonCleanupSheets_v671_ !== 'function') {
  throw new Error('v6.71 안전 시트 정리 함수를 찾지 못했습니다.');
}

var LOTTEON_CLEANUP_BUILD_ORIGINAL_V673 = buildLotteonSheetCleanupList_v671_;
var LOTTEON_CLEANUP_DELETE_ORIGINAL_V673 = deleteCheckedLotteonCleanupSheets_v671_;

buildLotteonSheetCleanupList_v671_ = function(options) {
  var ss = SpreadsheetApp.getActive();
  var manager = ss && ss.getSheetByName(LOTTEON_SHEET_CLEANUP_MANAGER_V671);
  if (manager && manager.getMaxRows() > 0) {
    try { manager.showRows(1, manager.getMaxRows()); } catch (ignore) {}
  }
  return LOTTEON_CLEANUP_BUILD_ORIGINAL_V673(options);
};

deleteCheckedLotteonCleanupSheets_v671_ = function() {
  var result = LOTTEON_CLEANUP_DELETE_ORIGINAL_V673();
  var ss = SpreadsheetApp.getActive();
  var manager = ss && ss.getSheetByName(LOTTEON_SHEET_CLEANUP_MANAGER_V671);
  if (!manager || manager.getLastRow() < 2) return result;

  var count = manager.getLastRow() - 1;
  var results = manager.getRange(2, 9, count, 1).getDisplayValues();
  for (var i = 0; i < results.length; i++) {
    if (/^삭제 완료(?:\s|$)/.test(String(results[i][0] || '').trim())) {
      try { manager.hideRows(i + 2); } catch (ignore) {}
    }
  }
  SpreadsheetApp.flush();
  return result;
};
