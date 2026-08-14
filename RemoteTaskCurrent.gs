var LOTTEON_REMOTE_TASK={id:'ISSUE75-v1.3-20260814',title:'v6.74 안전회수 운영반영+기간별 동기화 v1.3',enabled:true,statusSheet:'ISSUE75_실행상태'};
var ISSUE75_V13_SOURCE='https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/a06db3d87118d2269d094a82c638857c26cb7eff/RemoteTaskCurrent.gs';
function runLotteonRemoteTaskStartRemote_(){return issue75V13Exec_();}
function runLotteonRemoteTaskContinueRemote_(){return issue75V13Exec_();}
function issue75V13Exec_(){
  var res=UrlFetchApp.fetch(ISSUE75_V13_SOURCE+'?ts='+new Date().getTime(),{method:'get',muteHttpExceptions:true,followRedirects:true});
  var code=res.getContentText();
  if(res.getResponseCode()!==200||!code)throw new Error('Issue75 v1.2 source fetch 실패 HTTP '+res.getResponseCode());
  function rep(oldText,newText,label){if(code.indexOf(oldText)<0)throw new Error('Issue75 v1.3 patch anchor 누락: '+label);code=code.replace(oldText,newText);}
  rep("var LOTTEON_REMOTE_TASK={id:'ISSUE75-v1.2-20260814',title:'v6.74 안전회수 운영반영+기간별 동기화 v1.2',enabled:true,statusSheet:'ISSUE75_실행상태'};","var LOTTEON_REMOTE_TASK={id:'ISSUE75-v1.3-20260814',title:'v6.74 안전회수 운영반영+기간별 동기화 v1.3',enabled:true,statusSheet:'ISSUE75_실행상태'};",'task');
  rep("var X75V='v1.2-ISSUE75-V674-GUARDED-PRODUCTION-APPLY-FORMAT-PREFLIGHT';","var X75V='v1.3-ISSUE75-V674-GUARDED-PRODUCTION-APPLY-CARD-FORMAT-SAFE';",'version');
  rep("x75removeFilter_(card);cardW=true;","var cardFormatsBefore=card.getDataRange().getNumberFormats();x75removeFilter_(card);cardW=true;",'card format capture');
  rep("var now=x75card_(card);x75guardFinal_(now);var cm=x75matrixDiff_(pv.values,card.getDataRange().getValues());var cmd=x75matrixDiff_(ss.getSheetByName(X75PV).getDataRange().getDisplayValues(),card.getDataRange().getDisplayValues());if(cm||cmd)throw new Error('운영카드/Issue74 preview 불일치 typed='+cm+' display='+cmd);","var now=x75card_(card);x75guardFinal_(now);var cm=x75matrixDiff_(pv.values,card.getDataRange().getValues());var cfmt=x75matrixDiff_(cardFormatsBefore,card.getDataRange().getNumberFormats());if(cm||cfmt)throw new Error('운영카드 검증 불일치 typed='+cm+' numberFormat='+cfmt);",'card compare');
  rep("['FORMAT_BEFORE_VALUES_PREFLIGHT_typed',pre.typed]","['카드numberFormat차이',cfmt],['FORMAT_BEFORE_VALUES_PREFLIGHT_typed',pre.typed]",'status metric');
  rep("Issue75 v1.2 PASS: MATCHED 835 / NO_MATCH 22 / 기간별 동기화 완료","Issue75 v1.3 PASS: MATCHED 835 / NO_MATCH 22 / 기간별 동기화 완료",'toast');
  return eval(code+'\n;x75run_();');
}
