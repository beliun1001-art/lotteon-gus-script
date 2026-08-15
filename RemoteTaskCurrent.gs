var LOTTEON_REMOTE_TASK={id:'ISSUE79-V4-PINNED-PROBE-v4.1',title:'Issue79 v4 고정실행 + bridge 진단',enabled:true,statusSheet:'ISSUE79_V4_실행확인'};
var ISSUE79_V4_PINNED_SOURCE='https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/327e007b7bfb4e751cf8141417421a0366c4ce27/RemoteTaskCurrent.gs';
var ISSUE79_V41='v4.1-ISSUE79-PINNED-RUNNER-PROBE';
function runLotteonRemoteTaskStartRemote_(){return issue79V41Exec_();}
function runLotteonRemoteTaskContinueRemote_(){return issue79V41Exec_();}
function issue79V41Probe_(ss,rows){
  var sh=ss.getSheetByName('ISSUE79_V4_실행확인')||ss.insertSheet('ISSUE79_V4_실행확인');
  sh.clearContents();
  sh.getRange(1,1,1,2).setValues([['항목','값']]);
  if(rows&&rows.length)sh.getRange(2,1,rows.length,2).setValues(rows);
  sh.setFrozenRows(1);sh.getRange(1,1,1,2).setFontWeight('bold');sh.autoResizeColumns(1,2);
}
function issue79V41Exec_(){
  var ss=SpreadsheetApp.getActive();
  var started=new Date().toISOString();
  issue79V41Probe_(ss,[['version',ISSUE79_V41],['상태','RUNNING'],['실행시작',started],['고정sourceCommit','327e007b7bfb4e751cf8141417421a0366c4ce27'],['고정source','RemoteTaskCurrent.gs v4']]);
  try{
    var url=ISSUE79_V4_PINNED_SOURCE+'?ts='+new Date().getTime();
    var res=UrlFetchApp.fetch(url,{method:'get',muteHttpExceptions:true,followRedirects:true});
    var status=res.getResponseCode(),code=res.getContentText('UTF-8');
    if(status!==200||!code)throw new Error('v4 고정 source fetch 실패 HTTP '+status);
    if(code.indexOf("ISSUE79-NOMATCH16-DEEP-v4")<0||code.indexOf('v4.0-ISSUE79-V2SAFE-PRESERVE-DEEP-HINT-READONLY')<0)throw new Error('v4 고정 source marker 불일치');
    var result=eval(code+'\n;issue79V4Exec_();');
    issue79V41Probe_(ss,[['version',ISSUE79_V41],['상태','PASS'],['실행시작',started],['실행완료',new Date().toISOString()],['고정sourceCommit','327e007b7bfb4e751cf8141417421a0366c4ce27'],['sourceHTTP',status],['sourceMarker','V4_OK'],['실행결과version',result&&result.version?result.version:''],['실행결과SAFE',result&&result.totalSafe!==undefined?result.totalSafe:''],['실행결과잔여',result&&result.remain!==undefined?result.remain:'']]);
    return result;
  }catch(e){
    issue79V41Probe_(ss,[['version',ISSUE79_V41],['상태','ERROR'],['실행시작',started],['실행완료',new Date().toISOString()],['고정sourceCommit','327e007b7bfb4e751cf8141417421a0366c4ce27'],['오류',String(e&&e.message?e.message:e)]]);
    throw e;
  }
}
