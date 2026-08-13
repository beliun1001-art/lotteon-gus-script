var LOTTEON_REMOTE_TASK={id:'ISSUE61-v1.1-20260813',title:'Issue59 백업에서 부가세 카드매칭검증 1,355건 완전 복구 v1.1',enabled:true,statusSheet:'ISSUE61_복구상태'};
var I61_VERSION='v1.1-ISSUE61-FORMAT-FIRST-BATCHED-RECOVERY';
var I61_PROD='부가세_카드매칭검증',I61_BACKUP='ISSUE59_백업_부가세카드매칭검증',I61_PREVIEW='ISSUE54_카드매칭전체PREVIEW',I61_VAT='부가세_신고자료',I61_PERIOD='부가세_기간별',I61_HISTORY='카드사용내역_붙여넣기',I61_MASTER='카드_마스터',I61_I60='ISSUE60_복구진단상태',I61_TEMP='ISSUE61_복구준비';
var I61_STATE_KEY='ISSUE61_V11_STATE',I61_BATCH=150,I61_HANDLER='runLotteonRemoteTaskContinue';

function runLotteonRemoteTaskStartRemote_(){
  var ss=SpreadsheetApp.getActive();if(!ss)throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var prior=i61Load_();
  if(prior&&prior.spreadsheetId===ss.getId()&&prior.stage&&prior.stage!=='DONE')return runLotteonRemoteTaskContinueRemote_();
  i61DeleteTriggers_();
  var st=i61Ensure_(ss,LOTTEON_REMOTE_TASK.statusSheet);
  i61Status_(st,'RUNNING','PRECHECK','format-first 복구준비 사전검증 시작',{});
  try{
    [I61_PROD,I61_BACKUP,I61_PREVIEW,I61_VAT,I61_PERIOD,I61_HISTORY,I61_MASTER,I61_I60].forEach(function(n){i61Need_(ss,n);});
    i61ExpectKv_(ss,I61_I60,{'버전':'v1.1-ISSUE60-STRUCTURE-RECOVERY-DIAGNOSTIC','상태':'PASS','판정':'MIXED_OR_UNKNOWN','현재_keyedRows':85,'백업_keyedRows':1355,'백업_old집계일치':'YES','현재/백업_overlap':85,'현재/백업_material행차이':0,'backupOnly':1270,'currentOnly_vs_backup':0});
    var backup=i61Need_(ss,I61_BACKUP),preview=i61Need_(ss,I61_PREVIEW),prod=i61Need_(ss,I61_PROD);
    var b=i61Stats_(backup);i61Assert_(b,{orders:1355,matched:810,nonCard:494,ambiguous:1,noMatch:50,purchase:54807644});if(b.dup)throw new Error('백업 주문키 중복 '+b.dup);
    var p=i61Stats_(preview);i61Assert_(p,{orders:1355,matched:808,nonCard:498,ambiguous:0,noMatch:49,purchase:105762969});
    var formulas=backup.getDataRange().getFormulas(),fc=0;formulas.forEach(function(r){r.forEach(function(v){if(v)fc++;});});if(fc)throw new Error('백업 수식셀 존재 '+fc);

    var rows=backup.getLastRow(),cols=backup.getLastColumn();
    var vals=backup.getRange(1,1,rows,cols).getValues(),fmts=backup.getRange(1,1,rows,cols).getNumberFormats();
    var temp=i61Ensure_(ss,I61_TEMP);temp.clear();i61Grid_(temp,rows,cols);
    temp.getRange(1,1,rows,cols).setNumberFormats(fmts);
    temp.getRange(1,1,rows,cols).setValues(vals);
    SpreadsheetApp.flush();
    var t=i61Stats_(temp);i61Assert_(t,{orders:1355,matched:810,nonCard:494,ambiguous:1,noMatch:50,purchase:54807644});
    var typed=i61TypedDiff_(backup,temp),display=i61DisplayDiff_(backup,temp),material=i61Material_(backup,temp);
    if(typed)throw new Error('format-first 복구준비 typed 차이 '+typed);
    if(display)throw new Error('format-first 복구준비 display 차이 '+display);
    if(material.overlap!==1355||material.diff||material.leftOnly||material.rightOnly)throw new Error('format-first 복구준비 material 불일치 '+JSON.stringify(material));

    var state={spreadsheetId:ss.getId(),stage:'WRITE',offset:0,rows:rows,cols:cols,startedAt:new Date().toISOString(),completedAt:'',protected:i61Protected_(ss)};
    i61Save_(state);
    i61Grid_(prod,rows,cols);
    var cr=Math.max(prod.getLastRow(),rows),cc=Math.max(prod.getLastColumn(),cols);if(cr&&cc)prod.getRange(1,1,cr,cc).clearContent();
    prod.getRange(1,1,rows,cols).setNumberFormats(fmts);
    SpreadsheetApp.flush();
    i61Status_(st,'RUNNING','WRITE','복구준비 exact PASS; 기존 운영 복구 시작',{target:rows-1,processed:0,prepTyped:typed,prepDisplay:display});
    return i61RunBatch_(ss,state);
  }catch(e){var msg=String(e&&e.message?e.message:e);i61Status_(st,'ERROR','FAILED','기존 운영 카드검증 복구 실패',{error:msg});i61DeleteTriggers_();throw e;}
}

function runLotteonRemoteTaskContinueRemote_(){
  var state=i61Load_();if(!state||!state.spreadsheetId)throw new Error('Issue61 v1.1 실행 상태가 없습니다.');
  var ss=SpreadsheetApp.openById(state.spreadsheetId);
  if(state.stage==='DONE')return{ok:true,done:true,reason:'ALREADY_DONE'};
  return i61RunBatch_(ss,state);
}

function i61RunBatch_(ss,state){
  var lock=LockService.getScriptLock();if(!lock.tryLock(5000))return{ok:false,reason:'LOCK_BUSY'};
  try{
    var st=i61Ensure_(ss,LOTTEON_REMOTE_TASK.statusSheet),backup=i61Need_(ss,I61_BACKUP),temp=i61Need_(ss,I61_TEMP),prod=i61Need_(ss,I61_PROD);
    i61CheckProtected_(ss,state.protected);
    if(state.stage!=='WRITE')throw new Error('알 수 없는 복구 단계 '+state.stage);
    var count=Math.min(I61_BATCH,state.rows-state.offset);
    if(count>0){
      var src=temp.getRange(state.offset+1,1,count,state.cols),dst=prod.getRange(state.offset+1,1,count,state.cols);
      dst.setNumberFormats(src.getNumberFormats());
      dst.setValues(src.getValues());
      SpreadsheetApp.flush();
      var td=i61RangeTypedDiff_(src,dst),dd=i61RangeDisplayDiff_(src,dst);
      if(td||dd)throw new Error('운영 복구 배치 차이 offset='+state.offset+' typed='+td+' display='+dd);
      state.offset+=count;i61Save_(state);
    }
    if(state.offset<state.rows){
      i61Status_(st,'RUNNING','WRITE','기존 운영 카드검증 배치 복구 진행',{target:state.rows-1,processed:Math.max(0,state.offset-1)});
      i61Schedule_();return{ok:true,done:false,offset:state.offset};
    }

    prod.setFrozenRows(backup.getFrozenRows());prod.setFrozenColumns(backup.getFrozenColumns());for(var c=1;c<=state.cols;c++){try{prod.setColumnWidth(c,backup.getColumnWidth(c));}catch(ignore){}}
    var f=i61Stats_(prod);i61Assert_(f,{orders:1355,matched:810,nonCard:494,ambiguous:1,noMatch:50,purchase:54807644});if(f.dup)throw new Error('복구 운영 주문키 중복 '+f.dup);
    var md=i61Material_(backup,prod),typedAll=i61TypedDiff_(backup,prod),displayAll=i61DisplayDiff_(backup,prod);
    if(md.overlap!==1355||md.diff||md.leftOnly||md.rightOnly)throw new Error('복구 운영 material 불일치 '+JSON.stringify(md));
    if(typedAll)throw new Error('복구 운영/백업 typed 차이 '+typedAll);
    if(displayAll)throw new Error('복구 운영/백업 display 차이 '+displayAll);
    i61CheckProtected_(ss,state.protected);
    state.stage='DONE';state.completedAt=new Date().toISOString();i61Save_(state);i61DeleteTriggers_();
    i61Status_(st,'PASS','DONE','Issue59 정상 백업에서 기존 운영 카드검증 1,355건 완전 복구 완료',{orders:f.orders,matched:f.matched,nonCard:f.nonCard,ambiguous:f.ambiguous,noMatch:f.noMatch,purchase:f.purchase,overlap:md.overlap,materialDiff:md.diff,typedDiff:typedAll,displayDiff:displayAll,target:1355,processed:1355});
    PropertiesService.getScriptProperties().deleteProperty(I61_STATE_KEY);
    return{ok:true,done:true};
  }catch(e){var msg=String(e&&e.message?e.message:e),st2=i61Ensure_(ss,LOTTEON_REMOTE_TASK.statusSheet);i61Status_(st2,'ERROR','FAILED','기존 운영 카드검증 복구 실패',{processed:Math.max(0,(state.offset||0)-1),target:Math.max(0,(state.rows||1)-1),error:msg});i61DeleteTriggers_();throw e;
  }finally{lock.releaseLock();}
}

function i61Stats_(sh){var v=sh.getDataRange().getValues(),h=v[0].map(i61Text_),ia=i61Find_(h,['쿠팡계정ID']),io=i61Find_(h,['주문번호']),ip=i61Find_(h,['주문매입금액','매입금액']),is=i61Find_(h,['카드매칭상태']);if(ia<0||io<0||ip<0||is<0)throw new Error(sh.getName()+' 필수 헤더 누락');var o={orders:0,matched:0,nonCard:0,ambiguous:0,noMatch:0,purchase:0,dup:0,keys:{}};for(var r=1;r<v.length;r++){var k=i61Key_(v[r][ia],v[r][io]);if(!k)continue;if(o.keys[k])o.dup++;o.keys[k]=true;o.orders++;o.purchase+=i61Num_(v[r][ip]);var s=i61Text_(v[r][is]);if(s==='MATCHED'||s==='MASTER_MATCHED')o.matched++;else if(s==='NON_CARD')o.nonCard++;else if(s==='AMBIGUOUS')o.ambiguous++;else o.noMatch++;}return o;}
function i61Material_(a,b){var A=i61Map_(a),B=i61Map_(b),o={overlap:0,diff:0,leftOnly:0,rightOnly:0};Object.keys(A).forEach(function(k){if(B[k]===undefined)o.leftOnly++;else{o.overlap++;if(A[k]!==B[k])o.diff++;}});Object.keys(B).forEach(function(k){if(A[k]===undefined)o.rightOnly++;});return o;}
function i61Map_(sh){var v=sh.getDataRange().getValues(),h=v[0].map(i61Text_),ia=i61Find_(h,['쿠팡계정ID']),io=i61Find_(h,['주문번호']);if(ia<0||io<0)throw new Error(sh.getName()+' key 헤더 누락');var o={};for(var r=1;r<v.length;r++){var k=i61Key_(v[r][ia],v[r][io]);if(!k)continue;o[k]=v[r].map(i61Cell_).join('\u001f');}return o;}
function i61Protected_(ss){return{backup:i61Sig_(i61Need_(ss,I61_BACKUP)),preview:i61Sig_(i61Need_(ss,I61_PREVIEW)),vat:i61Sig_(i61Need_(ss,I61_VAT)),period:i61Sig_(i61Need_(ss,I61_PERIOD)),history:i61Sig_(i61Need_(ss,I61_HISTORY)),master:i61Sig_(i61Need_(ss,I61_MASTER))};}
function i61CheckProtected_(ss,b){var n=i61Protected_(ss);Object.keys(b).forEach(function(k){if(b[k]!==n[k])throw new Error('보호시트 변경 '+k);});}
function i61Sig_(sh){var v=sh.getDataRange().getValues(),h=2166136261;for(var r=0;r<v.length;r++)for(var c=0;c<v[r].length;c++){var s=i61Cell_(v[r][c])+'\u001f';for(var i=0;i<s.length;i++){h^=s.charCodeAt(i);h=Math.imul(h,16777619);}}return sh.getLastRow()+'x'+sh.getLastColumn()+'|'+(h>>>0).toString(16);}
function i61TypedDiff_(a,b){if(a.getLastRow()!==b.getLastRow()||a.getLastColumn()!==b.getLastColumn())return 999999;return i61RangeTypedDiff_(a.getDataRange(),b.getDataRange());}
function i61DisplayDiff_(a,b){if(a.getLastRow()!==b.getLastRow()||a.getLastColumn()!==b.getLastColumn())return 999999;return i61RangeDisplayDiff_(a.getDataRange(),b.getDataRange());}
function i61RangeTypedDiff_(a,b){var A=a.getValues(),B=b.getValues(),d=0;if(A.length!==B.length)return 999999;for(var r=0;r<A.length;r++){if(A[r].length!==B[r].length)return 999999;for(var c=0;c<A[r].length;c++)if(i61Cell_(A[r][c])!==i61Cell_(B[r][c]))d++;}return d;}
function i61RangeDisplayDiff_(a,b){var A=a.getDisplayValues(),B=b.getDisplayValues(),d=0;if(A.length!==B.length)return 999999;for(var r=0;r<A.length;r++){if(A[r].length!==B[r].length)return 999999;for(var c=0;c<A[r].length;c++)if(String(A[r][c])!==String(B[r][c]))d++;}return d;}
function i61Cell_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return'D:'+v.toISOString();if(typeof v==='number')return'N:'+String(v);if(typeof v==='boolean')return'B:'+String(v);return'T:'+i61Text_(v);}
function i61ExpectKv_(ss,n,e){var sh=i61Need_(ss,n),kv={};sh.getRange(1,1,sh.getLastRow(),Math.min(2,sh.getLastColumn())).getValues().forEach(function(r){var k=i61Text_(r[0]);if(k)kv[k]=r[1];});Object.keys(e).forEach(function(k){var w=e[k],a=kv[k];if(typeof w==='number'){if(Math.round(i61Num_(a))!==w)throw new Error(n+' '+k+' 불일치 '+a);}else if(i61Text_(a)!==String(w))throw new Error(n+' '+k+' 불일치 '+a);});}
function i61Assert_(a,e){Object.keys(e).forEach(function(k){if(Math.round(Number(a[k]||0))!==Math.round(Number(e[k]||0)))throw new Error(k+' 불일치 실제 '+a[k]+' 기대 '+e[k]);});}
function i61Status_(sh,status,stage,msg,x){x=x||{};var rows=[['항목','값'],['버전',I61_VERSION],['상태',status],['단계',stage],['메시지',msg],['처리행',x.processed||0],['대상행',x.target||0],['운영주문',x.orders||0],['MATCHED',x.matched||0],['NON_CARD',x.nonCard||0],['AMBIGUOUS',x.ambiguous||0],['NO_MATCH',x.noMatch||0],['주문매입금액합계',x.purchase||0],['운영/백업_overlap',x.overlap||0],['material행차이',x.materialDiff||0],['typed셀차이',x.typedDiff||0],['display셀차이',x.displayDiff||0],['복구준비_typed셀차이',x.prepTyped||0],['복구준비_display셀차이',x.prepDisplay||0],['Issue59백업 변경','0'],['Issue54preview 변경','0'],['부가세_신고자료 변경','0'],['부가세_기간별 변경','0'],['카드사용내역_붙여넣기 변경','0'],['카드_마스터 변경','0'],['오류',x.error||''],['완료시각',(status==='PASS'||status==='ERROR')?new Date().toISOString():''],['갱신시각',new Date().toISOString()]];sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.setFrozenRows(1);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setColumnWidth(1,240);sh.setColumnWidth(2,680);}
function i61Save_(s){PropertiesService.getScriptProperties().setProperty(I61_STATE_KEY,JSON.stringify(s));}function i61Load_(){var r=PropertiesService.getScriptProperties().getProperty(I61_STATE_KEY);return r?JSON.parse(r):null;}
function i61Schedule_(){i61DeleteTriggers_();ScriptApp.newTrigger(I61_HANDLER).timeBased().after(60*1000).create();}
function i61DeleteTriggers_(){ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()===I61_HANDLER){try{ScriptApp.deleteTrigger(t);}catch(ignore){}}});}
function i61Need_(ss,n){var s=ss.getSheetByName(n);if(!s)throw new Error('필수 시트 없음 '+n);return s;}function i61Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}function i61Grid_(s,r,c){if(s.getMaxRows()<r)s.insertRowsAfter(s.getMaxRows(),r-s.getMaxRows());if(s.getMaxColumns()<c)s.insertColumnsAfter(s.getMaxColumns(),c-s.getMaxColumns());}function i61Find_(h,n){for(var i=0;i<n.length;i++){var x=h.indexOf(n[i]);if(x>=0)return x;}return-1;}function i61Text_(v){return String(v==null?'':v).trim();}function i61Num_(v){var n=Number(typeof v==='number'?v:i61Text_(v).replace(/[,원\s]/g,''));return isNaN(n)?0:n;}function i61Key_(a,o){a=i61Text_(a).toLowerCase();o=i61Text_(o).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');return a&&o?a+'|'+o:'';}
