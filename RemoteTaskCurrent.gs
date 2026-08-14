var LOTTEON_REMOTE_TASK={id:'ISSUE63-v1.1-20260814',title:'Issue62 첫 배치 139셀 변형 sandbox 진단 v1.1',enabled:true,statusSheet:'ISSUE63_진단상태'};
var I63_VERSION='v1.1-ISSUE63-FIRST-BATCH-SANDBOX-DIAGNOSTIC';
var I63_PROD='부가세_카드매칭검증',I63_PREVIEW='ISSUE54_카드매칭전체PREVIEW',I63_BACKUP='ISSUE59_백업_부가세카드매칭검증',I63_VAT='부가세_신고자료',I63_PERIOD='부가세_기간별',I63_HISTORY='카드사용내역_붙여넣기',I63_MASTER='카드_마스터',I63_I62='ISSUE62_재반영상태',I63_DETAIL='ISSUE63_진단상세';

function runLotteonRemoteTaskStartRemote_(){
  var ss=SpreadsheetApp.getActive();
  if(!ss)throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var st=i63Ensure_(ss,LOTTEON_REMOTE_TASK.statusSheet);
  i63Status_(st,'RUNNING','PRECHECK','Issue63 v1.1 sandbox 진단 사전검증 시작',{});
  var protectedNames=[I63_PROD,I63_PREVIEW,I63_BACKUP,I63_VAT,I63_PERIOD,I63_HISTORY,I63_MASTER];
  var before={},sandNames=[];
  protectedNames.forEach(function(n){before[n]=i63Sig_(i63Need_(ss,n));});
  try{
    i63ExpectKv_(ss,I63_I62,{
      '버전':'v1.0-ISSUE62-FORMAT-FIRST-CORRECTED-APPLY',
      '상태':'ROLLED_BACK','단계':'DONE','운영주문':1355,'MATCHED':810,'NON_CARD':494,
      'AMBIGUOUS':1,'NO_MATCH':50,'주문매입금액합계':54807644,'롤백':1,
      '오류':'운영 재반영 배치 차이 offset=0 typed=139 display=139'
    });
    var prod=i63Need_(ss,I63_PROD),preview=i63Need_(ss,I63_PREVIEW),backup=i63Need_(ss,I63_BACKUP);
    var ps=i63Stats_(prod);i63Assert_(ps,{orders:1355,matched:810,nonCard:494,ambiguous:1,noMatch:50,purchase:54807644});
    var bs=i63Stats_(backup);i63Assert_(bs,{orders:1355,matched:810,nonCard:494,ambiguous:1,noMatch:50,purchase:54807644});
    var prodBackupTyped=i63SheetTypedDiff_(prod,backup),prodBackupDisplay=i63SheetDisplayDiff_(prod,backup);
    if(prodBackupTyped!==0||prodBackupDisplay!==0)throw new Error('현재 운영/Issue59백업 불일치 typed='+prodBackupTyped+' display='+prodBackupDisplay);
    var ns=i63Stats_(preview);i63Assert_(ns,{orders:1355,matched:808,nonCard:498,ambiguous:0,noMatch:49,v669:1161,v670:81,purchase:105762969});

    var totalRows=preview.getLastRow(),cols=preview.getLastColumn(),batchRows=Math.min(150,totalRows);
    var src=preview.getRange(1,1,batchRows,cols),vals=src.getValues();
    var allFormats=preview.getRange(1,1,totalRows,cols).getNumberFormats();
    var scenarios=[
      {key:'AS_IS_CLONE',prep:function(sh){sh.getDataRange().clearContent();}},
      {key:'CLEAR_VALIDATIONS',prep:function(sh){sh.getDataRange().clearContent();sh.getDataRange().clearDataValidations();}},
      {key:'FULL_CLEAR',prep:function(sh){sh.clear();}}
    ];
    var results=[],details=[['시나리오','행','열','헤더','src_typed','dst_typed','src_display','dst_display','src_format','dst_format','validation','note']];
    scenarios.forEach(function(sc){
      var nm='ISSUE63_SANDBOX_'+sc.key;
      var stale=ss.getSheetByName(nm);if(stale)ss.deleteSheet(stale);
      var clone=prod.copyTo(ss).setName(nm);sandNames.push(nm);
      sc.prep(clone);
      i63Grid_(clone,totalRows,cols);
      clone.getRange(1,1,totalRows,cols).setNumberFormats(allFormats);
      var dst=clone.getRange(1,1,batchRows,cols);
      dst.setValues(vals);
      SpreadsheetApp.flush();
      var cmp=i63RangeCompare_(src,dst);
      results.push({key:sc.key,typed:cmp.typed,display:cmp.display,format:cmp.format,validationCount:cmp.validationCount,samples:cmp.samples});
      cmp.samples.forEach(function(x){details.push([sc.key,x.r,x.c,x.header,x.srcTyped,x.dstTyped,x.srcDisplay,x.dstDisplay,x.srcFmt,x.dstFmt,x.validation,x.note]);});
    });

    protectedNames.forEach(function(n){if(before[n]!==i63Sig_(i63Need_(ss,n)))throw new Error('보호시트 변경 '+n);});
    var a=results[0],v=results[1],f=results[2],verdict='UNRESOLVED';
    if(a.typed>0&&a.display>0&&v.typed===0&&v.display===0)verdict='DATA_VALIDATION_CAUSE';
    else if(a.typed>0&&v.typed>0&&f.typed===0&&f.display===0)verdict='RESIDUAL_CELL_METADATA_CAUSE';
    else if(a.typed===0&&a.display===0&&v.typed===0&&v.display===0&&f.typed===0&&f.display===0)verdict='NON_REPRODUCIBLE';
    else if(f.typed>0||f.display>0)verdict='SHEET_LEVEL_OR_OTHER_CAUSE';

    var dsh=i63Ensure_(ss,I63_DETAIL);dsh.clearContents();
    if(details.length)dsh.getRange(1,1,details.length,12).setValues(details);
    dsh.setFrozenRows(1);dsh.setColumnWidth(1,190);dsh.setColumnWidth(4,180);for(var c=5;c<=12;c++)dsh.setColumnWidth(c,260);
    i63Status_(st,'PASS','DONE','Issue62 첫 배치 sandbox 원인진단 완료',{verdict:verdict,a:a,v:v,f:f,prodBackupTyped:prodBackupTyped,prodBackupDisplay:prodBackupDisplay});
    return{ok:true,done:true,verdict:verdict,results:results};
  }catch(e){
    var msg=String(e&&e.message?e.message:e);
    try{i63Status_(st,'ERROR','FAILED','Issue63 v1.1 sandbox 진단 실패',{error:msg});}catch(ignore){}
    throw e;
  }finally{
    sandNames.forEach(function(n){try{var s=ss.getSheetByName(n);if(s)ss.deleteSheet(s);}catch(ignore){}});
  }
}
function runLotteonRemoteTaskContinueRemote_(){return runLotteonRemoteTaskStartRemote_();}

function i63RangeCompare_(src,dst){
  var A=src.getValues(),B=dst.getValues(),AD=src.getDisplayValues(),BD=dst.getDisplayValues(),AF=src.getNumberFormats(),BF=dst.getNumberFormats(),valid=dst.getDataValidations(),notes=dst.getNotes();
  var headers=AD[0]||[],o={typed:0,display:0,format:0,validationCount:0,samples:[]};
  for(var r=0;r<A.length;r++)for(var c=0;c<A[r].length;c++){
    var ta=i63Cell_(A[r][c]),tb=i63Cell_(B[r][c]),da=String(AD[r][c]),db=String(BD[r][c]),fa=String(AF[r][c]),fb=String(BF[r][c]);
    if(valid[r][c])o.validationCount++;
    if(ta!==tb)o.typed++;if(da!==db)o.display++;if(fa!==fb)o.format++;
    if((ta!==tb||da!==db||fa!==fb)&&o.samples.length<80)o.samples.push({r:r+1,c:c+1,header:String(headers[c]||''),srcTyped:ta,dstTyped:tb,srcDisplay:da,dstDisplay:db,srcFmt:fa,dstFmt:fb,validation:i63Validation_(valid[r][c]),note:String(notes[r][c]||'')});
  }
  return o;
}
function i63Validation_(v){
  if(!v)return'';
  try{
    var vals=v.getCriteriaValues().map(function(x){
      if(Object.prototype.toString.call(x)==='[object Date]'&&!isNaN(x.getTime()))return x.toISOString();
      try{return String(x);}catch(ignore){return'[value]';}
    });
    return String(v.getCriteriaType())+'|'+vals.join('||')+'|allowInvalid='+v.getAllowInvalid();
  }catch(e){return'VALIDATION_PRESENT';}
}
function i63Stats_(sh){
  var v=sh.getDataRange().getValues(),h=v[0].map(i63Text_),ia=i63Find_(h,['쿠팡계정ID']),io=i63Find_(h,['주문번호']),ip=i63Find_(h,['주문매입금액','매입금액']),is=i63Find_(h,['카드매칭상태']),i69=i63Find_(h,['v6.69 2차귀속']),i70=i63Find_(h,['v6.70 3차귀속']);
  if(ia<0||io<0||ip<0||is<0)throw new Error(sh.getName()+' 필수 헤더 누락');
  var o={orders:0,matched:0,nonCard:0,ambiguous:0,noMatch:0,v669:0,v670:0,purchase:0,keys:{},dup:0};
  for(var r=1;r<v.length;r++){
    var k=i63Key_(v[r][ia],v[r][io]);if(!k)continue;if(o.keys[k])o.dup++;o.keys[k]=true;o.orders++;o.purchase+=i63Num_(v[r][ip]);
    var s=i63Text_(v[r][is]);if(s==='MATCHED'||s==='MASTER_MATCHED')o.matched++;else if(s==='NON_CARD')o.nonCard++;else if(s==='AMBIGUOUS')o.ambiguous++;else o.noMatch++;
    if(i69>=0&&i63Text_(v[r][i69])==='Y')o.v669++;if(i70>=0&&i63Text_(v[r][i70])==='Y')o.v670++;
  }
  return o;
}
function i63Assert_(a,e){Object.keys(e).forEach(function(k){if(Math.round(Number(a[k]||0))!==Math.round(Number(e[k]||0)))throw new Error(k+' 불일치 실제 '+a[k]+' 기대 '+e[k]);});}
function i63SheetTypedDiff_(a,b){if(a.getLastRow()!==b.getLastRow()||a.getLastColumn()!==b.getLastColumn())return 999999;return i63RangeTypedDiff_(a.getDataRange(),b.getDataRange());}
function i63RangeTypedDiff_(a,b){var A=a.getValues(),B=b.getValues(),d=0;for(var r=0;r<A.length;r++)for(var c=0;c<A[r].length;c++)if(i63Cell_(A[r][c])!==i63Cell_(B[r][c]))d++;return d;}
function i63SheetDisplayDiff_(a,b){if(a.getLastRow()!==b.getLastRow()||a.getLastColumn()!==b.getLastColumn())return 999999;var A=a.getDataRange().getDisplayValues(),B=b.getDataRange().getDisplayValues(),d=0;for(var r=0;r<A.length;r++)for(var c=0;c<A[r].length;c++)if(String(A[r][c])!==String(B[r][c]))d++;return d;}
function i63Sig_(sh){var v=sh.getDataRange().getValues(),h=2166136261;for(var r=0;r<v.length;r++)for(var c=0;c<v[r].length;c++){var s=i63Cell_(v[r][c])+'\u001f';for(var i=0;i<s.length;i++){h^=s.charCodeAt(i);h=Math.imul(h,16777619);}}return sh.getLastRow()+'x'+sh.getLastColumn()+'|'+(h>>>0).toString(16);}
function i63Cell_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return'D:'+v.toISOString();if(typeof v==='number')return'N:'+String(v);if(typeof v==='boolean')return'B:'+String(v);return'T:'+i63Text_(v);}
function i63ExpectKv_(ss,n,e){var sh=i63Need_(ss,n),kv={};sh.getRange(1,1,sh.getLastRow(),Math.min(2,sh.getLastColumn())).getValues().forEach(function(r){var k=i63Text_(r[0]);if(k)kv[k]=r[1];});Object.keys(e).forEach(function(k){var w=e[k],a=kv[k];if(typeof w==='number'){if(Math.round(i63Num_(a))!==w)throw new Error(n+' '+k+' 불일치 '+a);}else if(i63Text_(a)!==String(w))throw new Error(n+' '+k+' 불일치 '+a+' / 기대 '+w);});}
function i63Status_(sh,status,stage,msg,x){
  x=x||{};var a=x.a||{},v=x.v||{},f=x.f||{};
  var rows=[['항목','값'],['버전',I63_VERSION],['상태',status],['단계',stage],['메시지',msg],['판정',x.verdict||''],
    ['운영/백업_typed차이',x.prodBackupTyped||0],['운영/백업_display차이',x.prodBackupDisplay||0],
    ['AS_IS_CLONE_typed차이',a.typed||0],['AS_IS_CLONE_display차이',a.display||0],['AS_IS_CLONE_format차이',a.format||0],['AS_IS_CLONE_validation셀',a.validationCount||0],
    ['CLEAR_VALIDATIONS_typed차이',v.typed||0],['CLEAR_VALIDATIONS_display차이',v.display||0],['CLEAR_VALIDATIONS_format차이',v.format||0],['CLEAR_VALIDATIONS_validation셀',v.validationCount||0],
    ['FULL_CLEAR_typed차이',f.typed||0],['FULL_CLEAR_display차이',f.display||0],['FULL_CLEAR_format차이',f.format||0],['FULL_CLEAR_validation셀',f.validationCount||0],
    ['부가세_카드매칭검증 변경',0],['Issue54preview 변경',0],['Issue59백업 변경',0],['부가세_신고자료 변경',0],['부가세_기간별 변경',0],['카드사용내역_붙여넣기 변경',0],['카드_마스터 변경',0],['오류',x.error||''],['완료시각',(status==='PASS'||status==='ERROR')?new Date().toISOString():''],['갱신시각',new Date().toISOString()]];
  sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.setFrozenRows(1);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setColumnWidth(1,280);sh.setColumnWidth(2,760);
}
function i63Need_(ss,n){var s=ss.getSheetByName(n);if(!s)throw new Error('필수 시트 없음 '+n);return s;}
function i63Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}
function i63Grid_(s,r,c){if(s.getMaxRows()<r)s.insertRowsAfter(s.getMaxRows(),r-s.getMaxRows());if(s.getMaxColumns()<c)s.insertColumnsAfter(s.getMaxColumns(),c-s.getMaxColumns());}
function i63Find_(h,n){for(var i=0;i<n.length;i++){var x=h.indexOf(n[i]);if(x>=0)return x;}return-1;}
function i63Text_(v){return String(v==null?'':v).trim();}
function i63Num_(v){var n=Number(typeof v==='number'?v:i63Text_(v).replace(/[,원\s]/g,''));return isNaN(n)?0:n;}
function i63Key_(a,o){a=i63Text_(a).toLowerCase();o=i63Text_(o).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');return a&&o?a+'|'+o:'';}