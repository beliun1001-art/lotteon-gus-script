var LOTTEON_REMOTE_TASK={id:'ISSUE62-v1.0-20260814',title:'corrected VAT 카드매칭검증 format-first 배치 운영 재반영',enabled:true,statusSheet:'ISSUE62_재반영상태'};
var I62_VERSION='v1.0-ISSUE62-FORMAT-FIRST-CORRECTED-APPLY';
var I62_PROD='부가세_카드매칭검증',I62_OLD_BACKUP='ISSUE59_백업_부가세카드매칭검증',I62_PREVIEW='ISSUE54_카드매칭전체PREVIEW',I62_PREVIEW_STATUS='ISSUE54_실행상태',I62_I55='ISSUE55_진단상태',I62_I56='ISSUE56_판정상태',I62_I57='ISSUE57_진단상태',I62_I58='ISSUE58_진단상태',I62_I61='ISSUE61_복구상태',I62_VAT='부가세_신고자료',I62_PERIOD='부가세_기간별',I62_HISTORY='카드사용내역_붙여넣기',I62_MASTER='카드_마스터',I62_PREP='ISSUE62_재반영준비';
var I62_STATE_KEY='ISSUE62_V10_STATE',I62_BATCH=150,I62_HANDLER='runLotteonRemoteTaskContinue';

function runLotteonRemoteTaskStartRemote_(){
  var ss=SpreadsheetApp.getActive();if(!ss)throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var statusSh=i62Ensure_(ss,LOTTEON_REMOTE_TASK.statusSheet);
  var statusKv=i62Kv_(statusSh);
  if(i62Text_(statusKv['상태'])==='PASS'){
    var prodDone=ss.getSheetByName(I62_PROD),previewDone=ss.getSheetByName(I62_PREVIEW);
    if(prodDone&&previewDone&&i62TypedDiff_(prodDone,previewDone)===0&&i62DisplayDiff_(prodDone,previewDone)===0)return{ok:true,done:true,reason:'ALREADY_DONE'};
    throw new Error('Issue62 상태는 PASS이나 운영/preview가 다릅니다. 자동 재실행 금지.');
  }
  var prior=i62Load_();
  if(prior&&prior.spreadsheetId===ss.getId()&&prior.stage&&prior.stage!=='DONE')return runLotteonRemoteTaskContinueRemote_();
  i62DeleteTriggers_();
  i62Status_(statusSh,'RUNNING','PRECHECK','corrected preview 재반영 사전검증 시작',{});
  try{
    [I62_PROD,I62_OLD_BACKUP,I62_PREVIEW,I62_PREVIEW_STATUS,I62_I55,I62_I56,I62_I57,I62_I58,I62_I61,I62_VAT,I62_PERIOD,I62_HISTORY,I62_MASTER].forEach(function(n){i62Need_(ss,n);});
    i62ExpectKv_(ss,I62_I61,{'버전':'v1.1-ISSUE61-FORMAT-FIRST-BATCHED-RECOVERY','상태':'PASS','단계':'DONE','운영주문':1355,'MATCHED':810,'NON_CARD':494,'AMBIGUOUS':1,'NO_MATCH':50,'주문매입금액합계':54807644,'운영/백업_overlap':1355,'material행차이':0,'typed셀차이':0,'display셀차이':0,'복구준비_typed셀차이':0,'복구준비_display셀차이':0});
    i62ExpectKv_(ss,I62_PREVIEW_STATUS,{'버전':'v1.1-ISSUE54-CORRECTED-CARD-REMATCH-PREVIEW-REBUILD','상태':'PASS','preview주문':1355,'MATCHED':808,'NON_CARD':498,'AMBIGUOUS':0,'NO_MATCH':49,'v6.69 2차귀속':1161,'v6.70 3차귀속':81,'주문매입금액합계':105762969});
    i62ExpectKv_(ss,I62_I55,{'버전':'v1.0-ISSUE55-CARD-DELTA-DIAGNOSTIC','상태':'PASS','상태변경주문':12,'상태동일주문':1343});
    i62ExpectKv_(ss,I62_I56,{'버전':'v1.0-ISSUE56-CHANGED12-EVIDENCE-ADJUDICATION','상태':'PASS','상태변경주문':12,'AUTO_SAFE':3,'REVIEW_REQUIRED':5,'INVALID':4});
    i62ExpectKv_(ss,I62_I57,{'버전':'v1.0-ISSUE57-BLOCKED9-DEEP-DIAGNOSTIC','상태':'PASS','진단대상':9,'EXPLAINED_SAFE':5,'LIKELY_MATCHER_BUG':4,'DATA_GAP_REVIEW':0,'INVALID_STATE':0});
    i62ExpectKv_(ss,I62_I58,{'버전':'v1.0-ISSUE58-PAYMENT-SOURCE-CAUSE-SPLIT','상태':'PASS','진단대상':4,'DIAGNOSTIC_FALSE_POSITIVE':4,'UNRESOLVED':0});

    var prod=i62Need_(ss,I62_PROD),oldBackup=i62Need_(ss,I62_OLD_BACKUP),preview=i62Need_(ss,I62_PREVIEW);
    var oldStats=i62Stats_(prod);i62Assert_(oldStats,{orders:1355,matched:810,nonCard:494,ambiguous:1,noMatch:50,purchase:54807644});if(oldStats.dup)throw new Error('현재 운영 주문키 중복 '+oldStats.dup);
    var backupStats=i62Stats_(oldBackup);i62Assert_(backupStats,{orders:1355,matched:810,nonCard:494,ambiguous:1,noMatch:50,purchase:54807644});if(backupStats.dup)throw new Error('Issue59 백업 주문키 중복 '+backupStats.dup);
    var oldCmp=i62Material_(prod,oldBackup);if(oldCmp.overlap!==1355||oldCmp.diff||oldCmp.leftOnly||oldCmp.rightOnly)throw new Error('현재 운영/Issue59 백업 material 불일치 '+JSON.stringify(oldCmp));
    if(i62TypedDiff_(prod,oldBackup)!==0||i62DisplayDiff_(prod,oldBackup)!==0)throw new Error('현재 운영/Issue59 백업 typed/display 불일치');

    var newStats=i62Stats_(preview);i62Assert_(newStats,{orders:1355,matched:808,nonCard:498,ambiguous:0,noMatch:49,v669:1161,v670:81,purchase:105762969});if(newStats.dup)throw new Error('Issue54 preview 주문키 중복 '+newStats.dup);
    if(newStats.dupCanonical)throw new Error('Issue54 MATCHED canonical key 중복 '+newStats.dupCanonical);
    if(i62Overlap_(oldStats.keys,newStats.keys)!==1355)throw new Error('현재 운영/Issue54 preview normalized overlap 불일치');

    var formulas=preview.getDataRange().getFormulas(),fc=0;formulas.forEach(function(r){r.forEach(function(v){if(v)fc++;});});if(fc)throw new Error('Issue54 preview 수식셀 존재 '+fc);

    var before=i62Protected_(ss);
    var rows=preview.getLastRow(),cols=preview.getLastColumn(),vals=preview.getRange(1,1,rows,cols).getValues(),fmts=preview.getRange(1,1,rows,cols).getNumberFormats();
    var oldPrep=ss.getSheetByName(I62_PREP);if(oldPrep)ss.deleteSheet(oldPrep);
    var prep=ss.insertSheet(I62_PREP);i62Grid_(prep,rows,cols);
    prep.getRange(1,1,rows,cols).setNumberFormats(fmts);
    prep.getRange(1,1,rows,cols).setValues(vals);
    SpreadsheetApp.flush();
    var prepStats=i62Stats_(prep);i62Assert_(prepStats,{orders:1355,matched:808,nonCard:498,ambiguous:0,noMatch:49,v669:1161,v670:81,purchase:105762969});
    var prepMaterial=i62Material_(preview,prep),prepTyped=i62TypedDiff_(preview,prep),prepDisplay=i62DisplayDiff_(preview,prep);
    if(prepMaterial.overlap!==1355||prepMaterial.diff||prepMaterial.leftOnly||prepMaterial.rightOnly)throw new Error('재반영준비 material 불일치 '+JSON.stringify(prepMaterial));
    if(prepTyped)throw new Error('재반영준비/preview typed 차이 '+prepTyped);
    if(prepDisplay)throw new Error('재반영준비/preview display 차이 '+prepDisplay);
    i62CheckProtected_(ss,before);

    var state={spreadsheetId:ss.getId(),stage:'WRITE_INIT',offset:0,rows:rows,cols:cols,protectedBefore:before,prepTyped:prepTyped,prepDisplay:prepDisplay,startedAt:new Date().toISOString(),error:''};
    i62Save_(state);
    i62Status_(statusSh,'RUNNING','WRITE_INIT','재반영준비 검증 PASS; 운영 corrected 반영 준비',{processed:0,target:rows-1,prepTyped:prepTyped,prepDisplay:prepDisplay,rollback:'0'});
    return runLotteonRemoteTaskContinueRemote_();
  }catch(e){
    i62DeleteTriggers_();i62ClearState_();
    var msg=String(e&&e.message?e.message:e);
    i62Status_(statusSh,'ERROR','FAILED','corrected 운영 반영 전 사전검증 실패',{error:msg,rollback:'0'});
    throw e;
  }
}

function runLotteonRemoteTaskContinueRemote_(){
  var lock=LockService.getScriptLock();if(!lock.tryLock(5000))return{ok:false,reason:'LOCK_BUSY'};
  try{
    var state=i62Load_();if(!state||!state.spreadsheetId)throw new Error('Issue62 실행 상태가 없습니다.');
    var ss=SpreadsheetApp.openById(state.spreadsheetId),statusSh=i62Ensure_(ss,LOTTEON_REMOTE_TASK.statusSheet);
    try{
      if(state.stage==='WRITE_INIT')return i62WriteInit_(ss,statusSh,state);
      if(state.stage==='WRITE_BATCH')return i62WriteBatch_(ss,statusSh,state);
      if(state.stage==='VERIFY')return i62Verify_(ss,statusSh,state);
      if(state.stage==='ROLLBACK_INIT')return i62RollbackInit_(ss,statusSh,state);
      if(state.stage==='ROLLBACK_BATCH')return i62RollbackBatch_(ss,statusSh,state);
      if(state.stage==='ROLLBACK_VERIFY')return i62RollbackVerify_(ss,statusSh,state);
      if(state.stage==='DONE')return{ok:true,done:true};
      throw new Error('알 수 없는 Issue62 stage '+state.stage);
    }catch(e){
      var msg=String(e&&e.message?e.message:e);state.error=msg;
      if(state.stage.indexOf('ROLLBACK')===0){
        state.stage='DONE';i62Save_(state);i62DeleteTriggers_();
        i62Status_(statusSh,'ROLLBACK_ERROR','FAILED','corrected 반영 실패 후 롤백도 실패',{error:msg,rollback:'ERROR',processed:state.offset||0,target:(state.rows||1)-1});
        return{ok:false,done:true,status:'ROLLBACK_ERROR',error:msg};
      }
      if(state.stage==='WRITE_INIT'||state.stage==='WRITE_BATCH'||state.stage==='VERIFY'){
        state.stage='ROLLBACK_INIT';state.offset=0;i62Save_(state);
        i62Status_(statusSh,'ROLLBACK_PENDING','ROLLBACK_INIT','corrected 반영 오류; 기존 운영 자동 롤백 예약',{error:msg,rollback:'PENDING',processed:0,target:(state.rows||1)-1});
        i62Schedule_();return{ok:false,rollbackScheduled:true,error:msg};
      }
      state.stage='DONE';i62Save_(state);i62DeleteTriggers_();
      i62Status_(statusSh,'ERROR','FAILED','Issue62 실행 실패',{error:msg,rollback:'0'});return{ok:false,done:true,error:msg};
    }
  }finally{lock.releaseLock();}
}

function i62WriteInit_(ss,statusSh,state){
  var prod=i62Need_(ss,I62_PROD),oldBackup=i62Need_(ss,I62_OLD_BACKUP),prep=i62Need_(ss,I62_PREP);
  var oldStats=i62Stats_(prod);i62Assert_(oldStats,{orders:1355,matched:810,nonCard:494,ambiguous:1,noMatch:50,purchase:54807644});
  if(i62TypedDiff_(prod,oldBackup)!==0||i62DisplayDiff_(prod,oldBackup)!==0)throw new Error('WRITE_INIT 직전 운영/old backup 불일치');
  i62CheckProtected_(ss,state.protectedBefore);
  i62Grid_(prod,state.rows,state.cols);
  var clearRows=Math.max(prod.getLastRow(),state.rows),clearCols=Math.max(prod.getLastColumn(),state.cols);
  if(clearRows&&clearCols)prod.getRange(1,1,clearRows,clearCols).clearContent();
  prod.getRange(1,1,state.rows,state.cols).setNumberFormats(prep.getRange(1,1,state.rows,state.cols).getNumberFormats());
  SpreadsheetApp.flush();
  state.stage='WRITE_BATCH';state.offset=0;i62Save_(state);
  i62Status_(statusSh,'RUNNING','WRITE_BATCH','운영 contents 초기화 및 format-first 적용 완료',{processed:0,target:state.rows-1,prepTyped:state.prepTyped,prepDisplay:state.prepDisplay,rollback:'0'});
  return i62WriteBatch_(ss,statusSh,state);
}

function i62WriteBatch_(ss,statusSh,state){
  var prep=i62Need_(ss,I62_PREP),prod=i62Need_(ss,I62_PROD);
  var count=Math.min(I62_BATCH,state.rows-state.offset);
  if(count>0){
    var src=prep.getRange(state.offset+1,1,count,state.cols),dst=prod.getRange(state.offset+1,1,count,state.cols);
    dst.setNumberFormats(src.getNumberFormats());
    dst.setValues(src.getValues());
    SpreadsheetApp.flush();
    var td=i62RangeTypedDiff_(src,dst),dd=i62RangeDisplayDiff_(src,dst);if(td||dd)throw new Error('운영 재반영 배치 차이 offset='+state.offset+' typed='+td+' display='+dd);
    state.offset+=count;
  }
  if(state.offset>=state.rows){state.stage='VERIFY';i62Save_(state);i62Status_(statusSh,'RUNNING','VERIFY','1,355건 corrected 배치 쓰기 완료; 최종검증',{processed:state.rows-1,target:state.rows-1,prepTyped:state.prepTyped,prepDisplay:state.prepDisplay,rollback:'0'});return i62Verify_(ss,statusSh,state);}
  i62Save_(state);i62Status_(statusSh,'RUNNING','WRITE_BATCH','corrected 카드검증 배치 반영 중',{processed:Math.max(0,state.offset-1),target:state.rows-1,prepTyped:state.prepTyped,prepDisplay:state.prepDisplay,rollback:'0'});i62Schedule_();return{ok:true,stage:state.stage,offset:state.offset};
}

function i62Verify_(ss,statusSh,state){
  var prod=i62Need_(ss,I62_PROD),preview=i62Need_(ss,I62_PREVIEW);
  var f=i62Stats_(prod);i62Assert_(f,{orders:1355,matched:808,nonCard:498,ambiguous:0,noMatch:49,v669:1161,v670:81,purchase:105762969});if(f.dup)throw new Error('최종 운영 주문키 중복 '+f.dup);if(f.dupCanonical)throw new Error('최종 운영 MATCHED canonical key 중복 '+f.dupCanonical);
  var md=i62Material_(preview,prod),td=i62TypedDiff_(preview,prod),dd=i62DisplayDiff_(preview,prod);if(md.overlap!==1355||md.diff||md.leftOnly||md.rightOnly)throw new Error('최종 운영/preview material 불일치 '+JSON.stringify(md));if(td||dd)throw new Error('최종 운영/preview typed/display 불일치 typed='+td+' display='+dd);
  i62CheckProtected_(ss,state.protectedBefore);
  state.stage='DONE';i62Save_(state);i62DeleteTriggers_();i62ClearState_();
  i62Status_(statusSh,'PASS','DONE','corrected VAT 카드매칭검증 format-first 배치 운영 반영 완료',{processed:state.rows-1,target:state.rows-1,orders:f.orders,matched:f.matched,nonCard:f.nonCard,ambiguous:f.ambiguous,noMatch:f.noMatch,v669:f.v669,v670:f.v670,purchase:f.purchase,overlap:md.overlap,materialDiff:md.diff,typedDiff:td,displayDiff:dd,prepTyped:state.prepTyped,prepDisplay:state.prepDisplay,rollback:'0'});
  return{ok:true,done:true,status:'PASS'};
}

function i62RollbackInit_(ss,statusSh,state){
  var prod=i62Need_(ss,I62_PROD),backup=i62Need_(ss,I62_OLD_BACKUP);var rows=backup.getLastRow(),cols=backup.getLastColumn();i62Grid_(prod,rows,cols);
  var clearRows=Math.max(prod.getLastRow(),rows),clearCols=Math.max(prod.getLastColumn(),cols);if(clearRows&&clearCols)prod.getRange(1,1,clearRows,clearCols).clearContent();
  prod.getRange(1,1,rows,cols).setNumberFormats(backup.getRange(1,1,rows,cols).getNumberFormats());SpreadsheetApp.flush();
  state.stage='ROLLBACK_BATCH';state.offset=0;state.rollbackRows=rows;state.rollbackCols=cols;i62Save_(state);
  i62Status_(statusSh,'ROLLBACK_RUNNING','ROLLBACK_BATCH','기존 운영 format-first 자동 롤백 시작',{error:state.error,rollback:'1',processed:0,target:rows-1});
  return i62RollbackBatch_(ss,statusSh,state);
}

function i62RollbackBatch_(ss,statusSh,state){
  var backup=i62Need_(ss,I62_OLD_BACKUP),prod=i62Need_(ss,I62_PROD),count=Math.min(I62_BATCH,state.rollbackRows-state.offset);
  if(count>0){var src=backup.getRange(state.offset+1,1,count,state.rollbackCols),dst=prod.getRange(state.offset+1,1,count,state.rollbackCols);dst.setNumberFormats(src.getNumberFormats());dst.setValues(src.getValues());SpreadsheetApp.flush();var td=i62RangeTypedDiff_(src,dst),dd=i62RangeDisplayDiff_(src,dst);if(td||dd)throw new Error('롤백 배치 차이 offset='+state.offset+' typed='+td+' display='+dd);state.offset+=count;}
  if(state.offset>=state.rollbackRows){state.stage='ROLLBACK_VERIFY';i62Save_(state);return i62RollbackVerify_(ss,statusSh,state);}
  i62Save_(state);i62Status_(statusSh,'ROLLBACK_RUNNING','ROLLBACK_BATCH','기존 운영 자동 롤백 중',{error:state.error,rollback:'1',processed:Math.max(0,state.offset-1),target:state.rollbackRows-1});i62Schedule_();return{ok:true,stage:state.stage,offset:state.offset};
}

function i62RollbackVerify_(ss,statusSh,state){
  var backup=i62Need_(ss,I62_OLD_BACKUP),prod=i62Need_(ss,I62_PROD);var f=i62Stats_(prod);i62Assert_(f,{orders:1355,matched:810,nonCard:494,ambiguous:1,noMatch:50,purchase:54807644});var md=i62Material_(backup,prod),td=i62TypedDiff_(backup,prod),dd=i62DisplayDiff_(backup,prod);if(md.overlap!==1355||md.diff||md.leftOnly||md.rightOnly||td||dd)throw new Error('롤백 최종검증 불일치 material='+JSON.stringify(md)+' typed='+td+' display='+dd);i62CheckProtected_(ss,state.protectedBefore);
  var originalError=state.error||'';state.stage='DONE';i62Save_(state);i62DeleteTriggers_();i62ClearState_();i62Status_(statusSh,'ROLLED_BACK','DONE','corrected 반영 오류로 기존 운영 1,355건 자동 롤백 완료',{error:originalError,rollback:'1',processed:state.rollbackRows-1,target:state.rollbackRows-1,orders:f.orders,matched:f.matched,nonCard:f.nonCard,ambiguous:f.ambiguous,noMatch:f.noMatch,purchase:f.purchase,overlap:md.overlap,materialDiff:md.diff,typedDiff:td,displayDiff:dd});return{ok:false,done:true,status:'ROLLED_BACK',error:originalError};
}

function i62Stats_(sh){var v=sh.getDataRange().getValues(),h=v[0].map(i62Text_),ia=i62Find_(h,['쿠팡계정ID']),io=i62Find_(h,['주문번호']),ip=i62Find_(h,['주문매입금액','매입금액']),is=i62Find_(h,['카드매칭상태']),i669=i62Find_(h,['v6.69 2차귀속']),i670=i62Find_(h,['v6.70 3차귀속']),ic=i62Find_(h,['canonicalEvidenceKey']);if(ia<0||io<0||ip<0||is<0)throw new Error(sh.getName()+' 필수 헤더 누락');var o={orders:0,matched:0,nonCard:0,ambiguous:0,noMatch:0,v669:0,v670:0,purchase:0,dup:0,dupCanonical:0,keys:{}},canon={};for(var r=1;r<v.length;r++){var k=i62Key_(v[r][ia],v[r][io]);if(!k)continue;if(o.keys[k])o.dup++;o.keys[k]=true;o.orders++;o.purchase+=i62Num_(v[r][ip]);var s=i62Text_(v[r][is]);if(s==='MATCHED'||s==='MASTER_MATCHED')o.matched++;else if(s==='NON_CARD')o.nonCard++;else if(s==='AMBIGUOUS')o.ambiguous++;else o.noMatch++;if(i669>=0&&i62Text_(v[r][i669])==='Y')o.v669++;if(i670>=0&&i62Text_(v[r][i670])==='Y')o.v670++;if((s==='MATCHED'||s==='MASTER_MATCHED')&&ic>=0){var ck=i62Text_(v[r][ic]);if(ck){canon[ck]=(canon[ck]||0)+1;if(canon[ck]===2)o.dupCanonical++;}}}return o;}
function i62Material_(a,b){var A=i62Map_(a),B=i62Map_(b),o={overlap:0,diff:0,leftOnly:0,rightOnly:0};Object.keys(A).forEach(function(k){if(B[k]===undefined)o.leftOnly++;else{o.overlap++;if(A[k]!==B[k])o.diff++;}});Object.keys(B).forEach(function(k){if(A[k]===undefined)o.rightOnly++;});return o;}
function i62Map_(sh){var v=sh.getDataRange().getValues(),h=v[0].map(i62Text_),ia=i62Find_(h,['쿠팡계정ID']),io=i62Find_(h,['주문번호']);if(ia<0||io<0)throw new Error(sh.getName()+' key 헤더 누락');var o={};for(var r=1;r<v.length;r++){var k=i62Key_(v[r][ia],v[r][io]);if(!k)continue;o[k]=v[r].map(i62Cell_).join('\u001f');}return o;}
function i62Overlap_(a,b){var n=0;Object.keys(a||{}).forEach(function(k){if(b&&b[k])n++;});return n;}
function i62TypedDiff_(a,b){if(a.getLastRow()!==b.getLastRow()||a.getLastColumn()!==b.getLastColumn())return 999999;return i62RangeTypedDiff_(a.getDataRange(),b.getDataRange());}
function i62DisplayDiff_(a,b){if(a.getLastRow()!==b.getLastRow()||a.getLastColumn()!==b.getLastColumn())return 999999;return i62RangeDisplayDiff_(a.getDataRange(),b.getDataRange());}
function i62RangeTypedDiff_(a,b){var A=a.getValues(),B=b.getValues(),d=0;if(A.length!==B.length)return 999999;for(var r=0;r<A.length;r++){if(A[r].length!==B[r].length)return 999999;for(var c=0;c<A[r].length;c++)if(i62Cell_(A[r][c])!==i62Cell_(B[r][c]))d++;}return d;}
function i62RangeDisplayDiff_(a,b){var A=a.getDisplayValues(),B=b.getDisplayValues(),d=0;if(A.length!==B.length)return 999999;for(var r=0;r<A.length;r++){if(A[r].length!==B[r].length)return 999999;for(var c=0;c<A[r].length;c++)if(String(A[r][c])!==String(B[r][c]))d++;}return d;}
function i62Protected_(ss){return{oldBackup:i62Sig_(i62Need_(ss,I62_OLD_BACKUP)),preview:i62Sig_(i62Need_(ss,I62_PREVIEW)),vat:i62Sig_(i62Need_(ss,I62_VAT)),period:i62Sig_(i62Need_(ss,I62_PERIOD)),history:i62Sig_(i62Need_(ss,I62_HISTORY)),master:i62Sig_(i62Need_(ss,I62_MASTER)),i55:i62Sig_(i62Need_(ss,I62_I55)),i56:i62Sig_(i62Need_(ss,I62_I56)),i57:i62Sig_(i62Need_(ss,I62_I57)),i58:i62Sig_(i62Need_(ss,I62_I58)),i61:i62Sig_(i62Need_(ss,I62_I61))};}
function i62CheckProtected_(ss,b){var n=i62Protected_(ss);Object.keys(b||{}).forEach(function(k){if(b[k]!==n[k])throw new Error('보호시트 변경 '+k);});}
function i62Sig_(sh){var v=sh.getDataRange().getValues(),h=2166136261;for(var r=0;r<v.length;r++)for(var c=0;c<v[r].length;c++){var s=i62Cell_(v[r][c])+'\u001f';for(var i=0;i<s.length;i++){h^=s.charCodeAt(i);h=Math.imul(h,16777619);}}return sh.getLastRow()+'x'+sh.getLastColumn()+'|'+(h>>>0).toString(16);}
function i62Cell_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return'D:'+v.toISOString();if(typeof v==='number')return'N:'+String(v);if(typeof v==='boolean')return'B:'+String(v);return'T:'+i62Text_(v);}
function i62ExpectKv_(ss,n,e){var kv=i62Kv_(i62Need_(ss,n));Object.keys(e).forEach(function(k){var w=e[k],a=kv[k];if(typeof w==='number'){if(Math.round(i62Num_(a))!==w)throw new Error(n+' '+k+' 불일치 '+a+' / 기대 '+w);}else if(i62Text_(a)!==String(w))throw new Error(n+' '+k+' 불일치 '+a+' / 기대 '+w);});}
function i62Kv_(sh){var kv={};if(!sh||sh.getLastRow()<1)return kv;sh.getRange(1,1,sh.getLastRow(),Math.min(2,sh.getLastColumn())).getValues().forEach(function(r){var k=i62Text_(r[0]);if(k)kv[k]=r[1];});return kv;}
function i62Assert_(a,e){Object.keys(e).forEach(function(k){if(Math.round(Number(a[k]||0))!==Math.round(Number(e[k]||0)))throw new Error(k+' 불일치 실제 '+a[k]+' 기대 '+e[k]);});}
function i62Status_(sh,status,stage,msg,x){x=x||{};var rows=[['항목','값'],['버전',I62_VERSION],['상태',status],['단계',stage],['메시지',msg],['처리행',x.processed||0],['대상행',x.target||0],['운영주문',x.orders||0],['MATCHED',x.matched||0],['NON_CARD',x.nonCard||0],['AMBIGUOUS',x.ambiguous||0],['NO_MATCH',x.noMatch||0],['v6.69 2차귀속',x.v669||0],['v6.70 3차귀속',x.v670||0],['주문매입금액합계',x.purchase||0],['운영/preview_overlap',x.overlap||0],['material행차이',x.materialDiff||0],['typed셀차이',x.typedDiff||0],['display셀차이',x.displayDiff||0],['재반영준비_typed셀차이',x.prepTyped||0],['재반영준비_display셀차이',x.prepDisplay||0],['Issue59백업 변경','0'],['Issue54preview 변경','0'],['부가세_신고자료 변경','0'],['부가세_기간별 변경','0'],['카드사용내역_붙여넣기 변경','0'],['카드_마스터 변경','0'],['롤백',x.rollback||'0'],['오류',x.error||''],['완료시각',(status==='PASS'||status==='ERROR'||status==='ROLLED_BACK'||status==='ROLLBACK_ERROR')?new Date().toISOString():''],['갱신시각',new Date().toISOString()]];sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.setFrozenRows(1);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setColumnWidth(1,240);sh.setColumnWidth(2,680);}
function i62Schedule_(){i62DeleteTriggers_();ScriptApp.newTrigger(I62_HANDLER).timeBased().after(60*1000).create();}
function i62DeleteTriggers_(){ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()===I62_HANDLER){try{ScriptApp.deleteTrigger(t);}catch(ignore){}}});}
function i62Save_(s){PropertiesService.getScriptProperties().setProperty(I62_STATE_KEY,JSON.stringify(s));}
function i62Load_(){var raw=PropertiesService.getScriptProperties().getProperty(I62_STATE_KEY);return raw?JSON.parse(raw):null;}
function i62ClearState_(){PropertiesService.getScriptProperties().deleteProperty(I62_STATE_KEY);}
function i62Need_(ss,n){var s=ss.getSheetByName(n);if(!s)throw new Error('필수 시트 없음 '+n);return s;}
function i62Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}
function i62Grid_(s,r,c){if(s.getMaxRows()<r)s.insertRowsAfter(s.getMaxRows(),r-s.getMaxRows());if(s.getMaxColumns()<c)s.insertColumnsAfter(s.getMaxColumns(),c-s.getMaxColumns());}
function i62Find_(h,n){for(var i=0;i<n.length;i++){var x=h.indexOf(n[i]);if(x>=0)return x;}return-1;}
function i62Text_(v){return String(v==null?'':v).trim();}
function i62Num_(v){var n=Number(typeof v==='number'?v:i62Text_(v).replace(/[,원\s]/g,''));return isNaN(n)?0:n;}
function i62Key_(a,o){a=i62Text_(a).toLowerCase();o=i62Text_(o).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');return a&&o?a+'|'+o:'';}
