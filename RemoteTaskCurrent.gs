/**
 * Permanent remote task slot.
 * Current state: idle. The installed hourly bridge may continue to call this
 * file, but no operating task runs until this slot is replaced intentionally.
 */
const LOTTEON_REMOTE_TASK = {
  id: 'IDLE-v1-20260805',
  title: 'LOTTEON 원격 작업 대기',
  enabled: false
};

function runLotteonRemoteTaskStartRemote_() {
  return {
    ok: true,
    skipped: true,
    reason: 'NO_ACTIVE_REMOTE_TASK',
    taskId: LOTTEON_REMOTE_TASK.id
  };
}

function runLotteonRemoteTaskContinueRemote_() {
  return runLotteonRemoteTaskStartRemote_();
}
