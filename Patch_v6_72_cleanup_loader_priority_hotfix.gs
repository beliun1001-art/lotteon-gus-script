/** v6.72 Issue #34 hotfix: local loader calls the legacy cleanup entry first. */
var LOTTEON_PATCH_V672_CLEANUP_LOADER_PRIORITY_HOTFIX_LOADED = true;

/**
 * The installed loader resolves cleanup functions in this order:
 * cleanupOperationSheets_v653 -> showOperationSheetsOnly -> legacy fallbacks.
 * Override the first legacy entry so the existing menu reaches the v6.71
 * confirmation-based cleanup manager instead of hiding non-core sheets.
 */
function cleanupOperationSheets_v653() {
  if (typeof showOperationSheetsOnly !== 'function') {
    throw new Error('v6.71 안전 시트 정리 함수를 찾지 못했습니다.');
  }
  return showOperationSheetsOnly();
}
