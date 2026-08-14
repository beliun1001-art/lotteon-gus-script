var LOTTEON_REMOTE_TASK={id:'ISSUE65-v1.0-20260814',title:'Issue64 persistent clone metadata filter/table/dependency 원인분리',enabled:true,statusSheet:'ISSUE65_진단상태'};
var I65_VERSION='v1.0-ISSUE65-FILTER-TABLE-DEPENDENCY-SPLIT';
var I65_PROD='부가세_카드매칭검증',I65_PREVIEW='ISSUE54_카드매칭전체PREVIEW',I65_BACKUP='ISSUE59_백업_부가세카드매칭검증',I65_VAT='부가세_신고자료',I65_PERIOD='부가세_기간별',I65_HISTORY='카드사용내역_붙여넣기',I65_MASTER='카드_마스터',I65_I64='ISSUE64_진단상태',I65_DETAIL='ISSUE65_진단상세';

function runLotteonRemoteTaskStartRemote_(){
  var ss=SpreadsheetApp.getActive();
  if(!ss)throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var st=i65Ensure_(ss,LOTTEON_REMOTE_TASK.statusSheet);
  i65Status_(st,'RUNNING','PRECHECK','filter/table/dependency read-only sandbox 진단 시작',{});
  var protectedNames=[I65_PROD,I65_PREVIEW,I65_BACKUP,I65_VAT,I65_PERIOD,I65_HISTORY,I65_MASTER],before={},sand=[];
  protectedNames.forEach(function(n){before[n]=i65Sig_(i65Need_(ss,n));});
  try{
    i65ExpectKv_(ss,I65_I64,{
      '버전':'v1.0-ISSUE64-SHEET-LEVEL-CAUSE-SPLIT','상태':'PASS','단계':'DONE','판정':'PERSISTENT_CLONE_METADATA_CAUSE',
      'BASELINE_typed차이':139,'BASELINE_display차이':139,'UNMERGE_typed차이':139,'UNMERGE_display차이':139,
      'FRESH_typed차이':0,'FRESH_display차이':0,'부가세_카드매칭검증 변경':0
    });
    var prod=i65Need_(ss,I65_PROD),preview=i65Need_(ss,I65_PREVIEW),backup=i65Need_(ss,I65_BACKUP);
    var ps=i65Stats_(prod);i65Assert_(ps,{orders:1355,matched:810,nonCard:494,ambiguous:1,noMatch:50,purchase:54807644});
    var bs=i65Stats_(backup);i65Assert_(bs,{orders:1355,matched:810,nonCard:494,ambiguous:1,noMatch:50,purchase:54807644});
    if(i65SheetTypedDiff_(prod,backup)!==0||i65SheetDisplayDiff_(prod,backup)!==0)throw new Error('현재 운영/Issue59 백업 불일치');
    var ns=i65Stats_(preview);i65Assert_(ns,{orders:1355,matched:808,nonCard:498,ambiguous:0,noMatch:49,v669:1161,v670:81,purchase:105762969});

    var filterMeta=i65FilterMeta_(prod);
    var restMeta=i65RestSheetMeta_(ss,prod);
    var deps=i65Dependencies_(ss,prod);

    var rows=Math.min(150,preview.getLastRow()),cols=preview.getLastColumn(),src=preview.getRange(1,1,rows,cols),vals=src.getValues(),fmts=src.getNumberFormats();
    var details=[['구분','시나리오/시트','행','열','헤더/항목','값1','값2','추가정보']];
    i65AppendMeta_(details,filterMeta,restMeta,deps);

    var removed=i65ScenarioFilterRemoved_(ss,prod,'ISSUE65_SANDBOX_FILTER_REMOVED',src,vals,fmts,false,filterMeta);sand.push('ISSUE65_SANDBOX_FILTER_REMOVED');
    i65AppendCmp_(details,'FILTER_REMOVED',removed);
    var recreated=i65ScenarioFilterRemoved_(ss,prod,'ISSUE65_SANDBOX_FILTER_RECREATED',src,vals,fmts,true,filterMeta);sand.push('ISSUE65_SANDBOX_FILTER_RECREATED');
    i65AppendCmp_(details,'FILTER_RECREATED',recreated);
    var fresh=i65ScenarioFresh_(ss,'ISSUE65_SANDBOX_FRESH',rows,cols,src,vals,fmts);sand.push('ISSUE65_SANDBOX_FRESH');
    i65AppendCmp_(details,'FRESH',fresh);

    protectedNames.forEach(function(n){if(before[n]!==i65Sig_(i65Need_(ss,n)))throw new Error('보호시트 변경 '+n);});

    var verdict='UNRESOLVED';
    if(removed.typed===0&&removed.display===0){
      if(recreated.typed===0&&recreated.display===0&&recreated.filterRecreated==='YES')verdict='BASIC_FILTER_CAUSE_SAFE_RECREATE';
      else verdict='BASIC_FILTER_CAUSE_RECREATE_BREAKS';
    }else if(removed.typed>0&&fresh.typed===0&&fresh.display===0)verdict='HIDDEN_TABLE_OR_CLONE_METADATA';
    else if(removed.typed===0&&removed.display===0&&fresh.typed===0&&fresh.display===0&&!filterMeta.exists)verdict='NON_REPRODUCIBLE';

    var dsh=i65Ensure_(ss,I65_DETAIL);dsh.clearContents();
    if(details.length)dsh.getRange(1,1,details.length,8).setValues(details);
    dsh.setFrozenRows(1);dsh.setColumnWidth(1,180);dsh.setColumnWidth(2,260);dsh.setColumnWidth(5,220);dsh.setColumnWidth(6,360);dsh.setColumnWidth(7,360);dsh.setColumnWidth(8,520);

    i65Status_(st,'PASS','DONE','filter/table/dependency 원인분리 진단 완료',{verdict:verdict,filter:filterMeta,rest:restMeta,deps:deps,removed:removed,recreated:recreated,fresh:fresh});
    return{ok:true,done:true,verdict:verdict};
  }catch(e){
    var msg=String(e&&e.message?e.message:e);
    try{i65Status_(st,'ERROR','FAILED','Issue65 원인분리 진단 실패',{error:msg});}catch(ignore){}
    throw e;
  }finally{
    sand.forEach(function(n){try{var sh=ss.getSheetByName(n);if(sh)ss.deleteSheet(sh);}catch(ignore){}});
  }
}
function runLotteonRemoteTaskContinueRemote_(){return runLotteonRemoteTaskStartRemote_();}

function i65FilterMeta_(sh){
  var f=sh.getFilter(),o={exists:!!f,range:'',startRow:0,numRows:0,startCol:0,numCols:0,criteriaCount:0,criteria:{}};
  if(!f)return o;
  var r=f.getRange();o.range=r.getA1Notation();o.startRow=r.getRow();o.numRows=r.getNumRows();o.startCol=r.getColumn();o.numCols=r.getNumColumns();
  for(var c=o.startCol;c<o.startCol+o.numCols;c++){
    try{var cr=f.getColumnFilterCriteria(c);if(cr){o.criteriaCount++;o.criteria[c]=cr;}}catch(ignore){}
  }
  return o;
}
function i65ScenarioFilterRemoved_(ss,prod,name,src,vals,fmts,recreate,filterMeta){
  var stale=ss.getSheetByName(name);if(stale)ss.deleteSheet(stale);
  var sh=prod.copyTo(ss).setName(name),f=sh.getFilter();if(f)f.remove();
  sh.clear();i65Grid_(sh,Math.max(src.getNumRows(),prod.getMaxRows()),Math.max(src.getNumColumns(),prod.getMaxColumns()));
  var dst=sh.getRange(1,1,src.getNumRows(),src.getNumColumns());dst.setNumberFormats(fmts);dst.setValues(vals);SpreadsheetApp.flush();
  var filterRecreated='NO',criteriaRestored=0,filterError='';
  if(recreate&&filterMeta.exists){
    try{
      var rr=sh.getRange(filterMeta.startRow,filterMeta.startCol,Math.min(filterMeta.numRows,sh.getMaxRows()-filterMeta.startRow+1),Math.min(filterMeta.numCols,sh.getMaxColumns()-filterMeta.startCol+1));
      var nf=rr.createFilter();filterRecreated='YES';
      Object.keys(filterMeta.criteria).forEach(function(k){try{nf.setColumnFilterCriteria(Number(k),filterMeta.criteria[k]);criteriaRestored++;}catch(ignore){}});
      SpreadsheetApp.flush();
    }catch(e){filterError=String(e&&e.message?e.message:e);}
  }
  var cmp=i65RangeCompare_(src,dst);cmp.filterRecreated=filterRecreated;cmp.criteriaRestored=criteriaRestored;cmp.filterError=filterError;return cmp;
}
function i65ScenarioFresh_(ss,name,rows,cols,src,vals,fmts){
  var stale=ss.getSheetByName(name);if(stale)ss.deleteSheet(stale);
  var sh=ss.insertSheet(name);i65Grid_(sh,rows,cols);var dst=sh.getRange(1,1,rows,cols);dst.setNumberFormats(fmts);dst.setValues(vals);SpreadsheetApp.flush();return i65RangeCompare_(src,dst);
}
function i65RangeCompare_(src,dst){
  var A=src.getValues(),B=dst.getValues(),AD=src.getDisplayValues(),BD=dst.getDisplayValues(),AF=src.getNumberFormats(),BF=dst.getNumberFormats(),headers=AD[0]||[];
  var o={typed:0,display:0,format:0,samples:[]};
  for(var r=0;r<A.length;r++)for(var c=0;c<A[r].length;c++){
    var ta=i65Cell_(A[r][c]),tb=i65Cell_(B[r][c]),da=String(AD[r][c]),db=String(BD[r][c]),fa=String(AF[r][c]),fb=String(BF[r][c]);
    if(ta!==tb)o.typed++;if(da!==db)o.display++;if(fa!==fb)o.format++;
    if((ta!==tb||da!==db||fa!==fb)&&o.samples.length<60)o.samples.push({r:r+1,c:c+1,header:String(headers[c]||''),srcTyped:ta,dstTyped:tb,srcDisplay:da,dstDisplay:db,srcFmt:fa,dstFmt:fb});
  }
  return o;
}

function i65RestSheetMeta_(ss,prod){
  var o={status:'NOT_RUN',http:0,basicFilter:'',filterViews:-1,tables:-1,error:''};
  try{
    var url='https://sheets.googleapis.com/v4/spreadsheets/'+encodeURIComponent(ss.getId())+'?includeGridData=false';
    var res=UrlFetchApp.fetch(url,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});o.http=res.getResponseCode();
    if(o.http!==200){o.status='HTTP_'+o.http;o.error=res.getContentText().slice(0,500);return o;}
    var j=JSON.parse(res.getContentText()),sheets=j.sheets||[],target=null;
    for(var i=0;i<sheets.length;i++){if(sheets[i].properties&&sheets[i].properties.sheetId===prod.getSheetId()){target=sheets[i];break;}}
    if(!target){o.status='SHEET_NOT_FOUND';return o;}
    o.status='OK';o.basicFilter=target.basicFilter?'YES':'NO';o.filterViews=(target.filterViews||[]).length;o.tables=target.tables?target.tables.length:0;
  }catch(e){o.status='ERROR';o.error=String(e&&e.message?e.message:e);}
  return o;
}

function i65Dependencies_(ss,prod){
  var out={formulaRefs:0,namedRangeRefs:0,chartRefs:0,pivotRefs:0,formulaSamples:[],namedSamples:[],chartSamples:[],pivotSamples:[],errors:[]},targetId=prod.getSheetId();
  ss.getSheets().forEach(function(sh){
    if(sh.getName().indexOf('ISSUE65_SANDBOX_')===0)return;
    try{
      var hits=sh.createTextFinder(I65_PROD).matchFormulaText(true).findAll();
      hits.forEach(function(r){out.formulaRefs++;if(out.formulaSamples.length<30)out.formulaSamples.push(sh.getName()+'!'+r.getA1Notation()+'='+r.getFormula());});
    }catch(e){out.errors.push('formula:'+sh.getName()+':'+String(e&&e.message?e.message:e));}
    try{
      var charts=sh.getCharts();charts.forEach(function(ch,idx){try{var rs=ch.getRanges();for(var i=0;i<rs.length;i++){if(rs[i].getSheet().getSheetId()===targetId){out.chartRefs++;if(out.chartSamples.length<20)out.chartSamples.push(sh.getName()+'#chart'+idx+':'+rs[i].getA1Notation());break;}}}catch(ignore){}});
    }catch(e2){out.errors.push('chart:'+sh.getName()+':'+String(e2&&e2.message?e2.message:e2));}
    try{
      if(typeof sh.getPivotTables==='function')sh.getPivotTables().forEach(function(p,idx){try{var sr=p.getSourceDataRange();if(sr&&sr.getSheet().getSheetId()===targetId){out.pivotRefs++;if(out.pivotSamples.length<20)out.pivotSamples.push(sh.getName()+'#pivot'+idx+':'+sr.getA1Notation());}}catch(ignore){}});
    }catch(e3){out.errors.push('pivot:'+sh.getName()+':'+String(e3&&e3.message?e3.message:e3));}
  });
  try{ss.getNamedRanges().forEach(function(nr){try{var r=nr.getRange();if(r&&r.getSheet().getSheetId()===targetId){out.namedRangeRefs++;if(out.namedSamples.length<30)out.namedSamples.push(nr.getName()+':'+r.getA1Notation());}}catch(ignore){}});}catch(e4){out.errors.push('named:'+String(e4&&e4.message?e4.message:e4));}
  return out;
}

function i65AppendMeta_(d,f,r,x){
  d.push(['META','FILTER',0,0,'exists',f.exists?'YES':'NO',f.range,'criteria='+f.criteriaCount]);
  d.push(['META','REST',0,0,'status',r.status,'HTTP '+r.http,'basicFilter='+r.basicFilter+' filterViews='+r.filterViews+' tables='+r.tables+' err='+r.error]);
  d.push(['DEPENDENCY','FORMULA',0,0,'count',x.formulaRefs,'',x.formulaSamples.join('\n')]);
  d.push(['DEPENDENCY','NAMED_RANGE',0,0,'count',x.namedRangeRefs,'',x.namedSamples.join('\n')]);
  d.push(['DEPENDENCY','CHART',0,0,'count',x.chartRefs,'',x.chartSamples.join('\n')]);
  d.push(['DEPENDENCY','PIVOT',0,0,'count',x.pivotRefs,'',x.pivotSamples.join('\n')]);
  if(x.errors.length)d.push(['DEPENDENCY','ERRORS',0,0,'count',x.errors.length,'',x.errors.slice(0,30).join('\n')]);
}
function i65AppendCmp_(d,name,o){
  d.push(['COMPARE',name,0,0,'summary','typed='+o.typed+' display='+o.display+' format='+o.format,'','filterRecreated='+(o.filterRecreated||'')+' criteriaRestored='+(o.criteriaRestored||0)+' filterError='+(o.filterError||'')]);
  o.samples.forEach(function(s){d.push(['MISMATCH',name,s.r,s.c,s.header,s.srcTyped+' / '+s.srcDisplay,s.dstTyped+' / '+s.dstDisplay,'srcFmt='+s.srcFmt+' dstFmt='+s.dstFmt]);});
}

function i65Stats_(sh){
  var v=sh.getDataRange().getValues(),h=v[0].map(i65Text_),ia=i65Find_(h,['쿠팡계정ID']),io=i65Find_(h,['주문번호']),ip=i65Find_(h,['주문매입금액','매입금액']),is=i65Find_(h,['카드매칭상태']),i69=i65Find_(h,['v6.69 2차귀속']),i70=i65Find_(h,['v6.70 3차귀속']);
  if(ia<0||io<0||ip<0||is<0)throw new Error(sh.getName()+' 필수 헤더 누락');
  var o={orders:0,matched:0,nonCard:0,ambiguous:0,noMatch:0,v669:0,v670:0,purchase:0};
  for(var r=1;r<v.length;r++){var k=i65Key_(v[r][ia],v[r][io]);if(!k)continue;o.orders++;o.purchase+=i65Num_(v[r][ip]);var s=i65Text_(v[r][is]);if(s==='MATCHED'||s==='MASTER_MATCHED')o.matched++;else if(s==='NON_CARD')o.nonCard++;else if(s==='AMBIGUOUS')o.ambiguous++;else o.noMatch++;if(i69>=0&&i65Text_(v[r][i69])==='Y')o.v669++;if(i70>=0&&i65Text_(v[r][i70])==='Y')o.v670++;}
  return o;
}
function i65Assert_(a,e){Object.keys(e).forEach(function(k){if(Math.round(Number(a[k]||0))!==Math.round(Number(e[k]||0)))throw new Error(k+' 불일치 실제 '+a[k]+' 기대 '+e[k]);});}
function i65SheetTypedDiff_(a,b){if(a.getLastRow()!==b.getLastRow()||a.getLastColumn()!==b.getLastColumn())return 999999;var A=a.getDataRange().getValues(),B=b.getDataRange().getValues(),d=0;for(var r=0;r<A.length;r++)for(var c=0;c<A[r].length;c++)if(i65Cell_(A[r][c])!==i65Cell_(B[r][c]))d++;return d;}
function i65SheetDisplayDiff_(a,b){if(a.getLastRow()!==b.getLastRow()||a.getLastColumn()!==b.getLastColumn())return 999999;var A=a.getDataRange().getDisplayValues(),B=b.getDataRange().getDisplayValues(),d=0;for(var r=0;r<A.length;r++)for(var c=0;c<A[r].length;c++)if(String(A[r][c])!==String(B[r][c]))d++;return d;}
function i65Sig_(sh){var v=sh.getDataRange().getValues(),h=2166136261;for(var r=0;r<v.length;r++)for(var c=0;c<v[r].length;c++){var s=i65Cell_(v[r][c])+'\u001f';for(var i=0;i<s.length;i++){h^=s.charCodeAt(i);h=Math.imul(h,16777619);}}return sh.getLastRow()+'x'+sh.getLastColumn()+'|'+(h>>>0).toString(16);}
function i65Cell_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return'D:'+v.toISOString();if(typeof v==='number')return'N:'+String(v);if(typeof v==='boolean')return'B:'+String(v);return'T:'+i65Text_(v);}
function i65ExpectKv_(ss,n,e){var sh=i65Need_(ss,n),kv={};sh.getRange(1,1,sh.getLastRow(),Math.min(2,sh.getLastColumn())).getValues().forEach(function(r){var k=i65Text_(r[0]);if(k)kv[k]=r[1];});Object.keys(e).forEach(function(k){var w=e[k],a=kv[k];if(typeof w==='number'){if(Math.round(i65Num_(a))!==w)throw new Error(n+' '+k+' 불일치 '+a);}else if(i65Text_(a)!==String(w))throw new Error(n+' '+k+' 불일치 '+a+' / 기대 '+w);});}
function i65Status_(sh,status,stage,msg,x){
  x=x||{};var f=x.filter||{},r=x.rest||{},d=x.deps||{},a=x.removed||{},b=x.recreated||{},c=x.fresh||{};
  var rows=[['항목','값'],['버전',I65_VERSION],['상태',status],['단계',stage],['메시지',msg],['판정',x.verdict||''],
    ['운영_filter',f.exists?'YES':'NO'],['운영_filter범위',f.range||''],['운영_filterCriteria수',f.criteriaCount||0],
    ['REST상태',r.status||''],['REST_HTTP',r.http||0],['REST_basicFilter',r.basicFilter||''],['REST_filterViews',r.filterViews==null?'':r.filterViews],['REST_tables',r.tables==null?'':r.tables],
    ['formula참조수',d.formulaRefs||0],['namedRange참조수',d.namedRangeRefs||0],['chart참조수',d.chartRefs||0],['pivot참조수',d.pivotRefs||0],['dependency오류수',(d.errors||[]).length],
    ['FILTER_REMOVED_typed차이',a.typed||0],['FILTER_REMOVED_display차이',a.display||0],['FILTER_REMOVED_format차이',a.format||0],
    ['FILTER_RECREATED_typed차이',b.typed||0],['FILTER_RECREATED_display차이',b.display||0],['FILTER_RECREATED_format차이',b.format||0],['FILTER_RECREATED_filter',b.filterRecreated||''],['FILTER_RECREATED_criteria복원',b.criteriaRestored||0],['FILTER_RECREATED_오류',b.filterError||''],
    ['FRESH_typed차이',c.typed||0],['FRESH_display차이',c.display||0],['FRESH_format차이',c.format||0],
    ['부가세_카드매칭검증 변경',0],['Issue54preview 변경',0],['Issue59백업 변경',0],['부가세_신고자료 변경',0],['부가세_기간별 변경',0],['카드사용내역_붙여넣기 변경',0],['카드_마스터 변경',0],
    ['오류',x.error||''],['완료시각',(status==='PASS'||status==='ERROR')?new Date().toISOString():''],['갱신시각',new Date().toISOString()]];
  sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.setFrozenRows(1);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setColumnWidth(1,300);sh.setColumnWidth(2,780);
}
function i65Need_(ss,n){var s=ss.getSheetByName(n);if(!s)throw new Error('필수 시트 없음 '+n);return s;}
function i65Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}
function i65Grid_(s,r,c){if(s.getMaxRows()<r)s.insertRowsAfter(s.getMaxRows(),r-s.getMaxRows());if(s.getMaxColumns()<c)s.insertColumnsAfter(s.getMaxColumns(),c-s.getMaxColumns());}
function i65Find_(h,n){for(var i=0;i<n.length;i++){var x=h.indexOf(n[i]);if(x>=0)return x;}return-1;}
function i65Text_(v){return String(v==null?'':v).trim();}
function i65Num_(v){var n=Number(typeof v==='number'?v:i65Text_(v).replace(/[,원\s]/g,''));return isNaN(n)?0:n;}
function i65Key_(a,o){a=i65Text_(a).toLowerCase();o=i65Text_(o).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');return a&&o?a+'|'+o:'';}
