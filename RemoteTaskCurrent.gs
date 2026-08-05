/**
 * Permanent remote task slot.
 * Current state: idle while Issue #43 waits for the separate payment-method
 * source collection work to be completed and reflected in production data.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'IDLE-ISSUE43-BLOCKED-20260805',
  title: 'Issue43 결제수단 상위작업 완료 대기',
  enabled: false
};

function runLotteonRemoteTaskStartRemote_() {
  return {
    ok: true,
    skipped: true,
    reason: 'BLOCKED_ON_PAYMENT_METHOD_SOURCE_WORK',
    taskId: LOTTEON_REMOTE_TASK.id
  };
}

function runLotteonRemoteTaskContinueRemote_() {
  return runLotteonRemoteTaskStartRemote_();
}
