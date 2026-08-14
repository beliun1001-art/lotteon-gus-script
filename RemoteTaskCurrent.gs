var LOTTEON_REMOTE_TASK={id:'ISSUE71-v1.0-20260814',title:'VAT corrected 임시 시트 최종 정리',enabled:true,statusSheet:''};
var I71_VERSION='v1.0-ISSUE71-TEMP-SHEET-CLEANUP';
var I71_KEEP=['매출데이터_붙여넣기','부가세_신고자료','부가세_카드매칭검증','부가세_기간별','카드사용내역_붙여넣기','카드_마스터'];
var I71_TARGET_RE=/^ISSUE(?:5[3-9]|6[0-8])(?:_|$)/i;

function runLotteonRemoteTaskStartRemote_(){
  var ss=SpreadsheetApp.getActive();
  var before={};
  var missing=[];
  I71_KEEP.forEach(function(name){
    var sh=ss.getSheetByName(name);
    if(!sh){missing.push(name);return;}
    before[name]=i71Sig_(sh);
  });
  if(missing.length) throw new Error('핵심 시트 누락으로 정리 중단: '+missing.join(', '));

  var candidates=ss.getSheets().filter(function(sh){
    var name=sh.getName();
    return I71_TARGET_RE.test(name) && I71_KEEP.indexOf(name)<0;
  }).map(function(sh){return sh.getName();});

  candidates.forEach(function(name){
    var sh=ss.getSheetByName(name);
    if(sh) ss.deleteSheet(sh);
  });
  SpreadsheetApp.flush();

  var changed=[];
  var missingAfter=[];
  I71_KEEP.forEach(function(name){
    var sh=ss.getSheetByName(name);
    if(!sh){missingAfter.push(name);return;}
    if(i71Sig_(sh)!==before[name]) changed.push(name);
  });
  if(missingAfter.length) throw new Error('정리 후 핵심 시트 누락: '+missingAfter.join(', '));
  if(changed.length) throw new Error('정리 후 핵심 시트 값 변경 감지: '+changed.join(', '));

  var remaining=ss.getSheets().map(function(sh){return sh.getName();}).filter(function(name){return I71_TARGET_RE.test(name);});
  if(remaining.length) throw new Error('임시 시트 잔여: '+remaining.join(', '));

  var result={
    ok:true,
    done:true,
    version:I71_VERSION,
    deletedCount:candidates.length,
    deletedSheets:candidates,
    remainingTempCount:0,
    coreChangedCount:0,
    coreSheets:I71_KEEP.slice(),
    completedAt:new Date().toISOString()
  };
  PropertiesService.getScriptProperties().setProperty('ISSUE71_LAST_RESULT',JSON.stringify(result));
  var msg='임시 시트 정리 완료: 삭제 '+candidates.length+'개 / 잔여 0 / 핵심변경 0';
  Logger.log(msg+' / '+JSON.stringify(candidates));
  try{ss.toast(msg,'LOTTEON',8);}catch(e){}
  return result;
}

function runLotteonRemoteTaskContinueRemote_(){return runLotteonRemoteTaskStartRemote_();}

function i71Sig_(sh){
  var rows=sh.getLastRow(), cols=sh.getLastColumn();
  if(rows<1||cols<1) return rows+'x'+cols+':EMPTY';
  var values=sh.getRange(1,1,rows,cols).getValues();
  var text=JSON.stringify(values,function(_k,v){
    if(Object.prototype.toString.call(v)==='[object Date]') return {__date__:v.getTime()};
    if(typeof v==='number' && isNaN(v)) return {__nan__:true};
    return v;
  });
  var bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,text,Utilities.Charset.UTF_8);
  var hex=bytes.map(function(b){var n=(b<0?b+256:b).toString(16);return n.length===1?'0'+n:n;}).join('');
  return rows+'x'+cols+':'+hex;
}
