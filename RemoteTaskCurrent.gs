var LOTTEON_REMOTE_TASK={
  id:'IDLE-20260814-AFTER-ISSUE69',
  title:'VAT corrected 작업 완료 - remote runner idle',
  enabled:false,
  statusSheet:'REMOTE_TASK_IDLE'
};

function runLotteonRemoteTaskStartRemote_(){
  return {ok:true,done:true,idle:true,message:'현재 예약된 원격 작업이 없습니다.'};
}

function runLotteonRemoteTaskContinueRemote_(){
  return runLotteonRemoteTaskStartRemote_();
}
