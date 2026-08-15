var LOTTEON_REMOTE_TASK={id:'ISSUE79-V441-COLUMNFIX',title:'Issue79 v4.4 주문번호대조 출력열수 수정',enabled:true,statusSheet:'ISSUE79_현재상태대조상태'};
var I79V441_SOURCE='https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/c80daf194aa718e71f2e233deb4c8b097dbb3b64/RemoteTaskCurrent.gs';
function runLotteonRemoteTaskStartRemote_(){return i79v441run_();}
function runLotteonRemoteTaskContinueRemote_(){return i79v441run_();}
function i79v441run_(){
  var res=UrlFetchApp.fetch(I79V441_SOURCE+'?ts='+new Date().getTime(),{method:'get',muteHttpExceptions:true,followRedirects:true});
  var src=res.getContentText('UTF-8');
  if(res.getResponseCode()!==200||!src)throw new Error('Issue79 v4.4 pinned source fetch 실패 HTTP '+res.getResponseCode());
  var marker="var I79V44='v4.4-ISSUE79-ORDERKEY-CURRENT-VAT-I74-READONLY';";
  if(src.indexOf(marker)<0)throw new Error('Issue79 v4.4 source marker 불일치');
  var bad="rows.push(['CURRENT_ONLY',c.orderNo,'',c.status,'',c.purchase,'',c.purchase,c.business,c.account,c.reason]);";
  var good="rows.push(['CURRENT_ONLY',c.orderNo,'',c.status,'',c.purchase,c.purchase,c.business,c.account,c.reason]);";
  if(src.indexOf(bad)<0)throw new Error('Issue79 v4.4 CURRENT_ONLY column anchor 누락');
  src=src.replace(bad,good);
  src=src.replace(marker,"var I79V44='v4.4.1-ISSUE79-ORDERKEY-CURRENT-VAT-I74-READONLY';");
  return eval(src+'\n;i79v44run_();');
}
