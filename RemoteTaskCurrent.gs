/**
 * Issue #60 v1.1 read-only structure/recovery diagnostic after Issue59 ROLLBACK_ERROR.
 * Never writes production/backup/preview/protected sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE60-v1.1-STRUCTURE-20260813',
  title: 'Issue59 ROLLBACK_ERROR 운영 카드검증 구조/복구상태 read-only 진단',
  enabled: true,
  statusSheet: 'ISSUE60_복구진단상태'
};

var I60_VERSION = 'v1.1-ISSUE60-STRUCTURE-RECOVERY-DIAGNOSTIC';
var I60_PROD = '부가세_카드매칭검증';
var I60_BACKUP = 'ISSUE59_백업_부가세카드매칭검증';
var I60_PREVIEW = 'ISSUE54_카드매칭전체PREVIEW';
var I60_VAT = '부가세_신고자료';
var I60_PERIOD = '부가세_기간별';
var I60_HISTORY = '카드사용내역_붙여넣기';
var I60_MASTER = '카드_마스터';
var I60_DETAIL = 'ISSUE60_복구진단상세';

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var statusSh = i60Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  i60WriteStatus_(statusSh, 'RUNNING', 'STRUCTURE_CHECK', '운영/백업/preview 구조 read-only 진단 시작', {});

  var names = [I60_PROD,I60_BACKUP,I60_PREVIEW,I60_VAT,I60_PERIOD,I60_HISTORY,I60_MASTER];
  var src = {};
  names.forEach(function(name){ src[name] = i60Need_(ss,name); });
  var before = {};
  names.forEach(function(name){ before[name]=i60SheetSigDisplay_(src[name]); });

  try {
    var prod = i60Inspect_(src[I60_PROD]);
    var backup = i60Inspect_(src[I60_BACKUP]);
    var preview = i60Inspect_(src[I60_PREVIEW]);

    var oldStats = i60IsOldStats_(prod.stats);
    var newStats = i60IsNewStats_(prod.stats);
    var backupOld = i60IsOldStats_(backup.stats);
    var previewNew = i60IsNewStats_(preview.stats);

    var cb = i60CompareByKey_(prod, backup);
    var cp = i60CompareByKey_(prod, preview);

    var verdict = 'MIXED_OR_UNKNOWN';
    if (oldStats && cb.overlap===1355 && cb.onlyLeft===0 && cb.onlyRight===0 && cb.materialDiff===0) verdict='LOGICALLY_OLD_ROLLBACK_OK';
    else if (newStats && cp.overlap===1355 && cp.onlyLeft===0 && cp.onlyRight===0 && cp.materialDiff===0) verdict='LOGICALLY_NEW_APPLY_PRESENT';
    else if (oldStats) verdict='OLD_STATS_BUT_CELL_DIFF';
    else if (newStats) verdict='NEW_STATS_BUT_CELL_DIFF';

    var detail = [];
    detail.push(['구분','시트','행','열','헤더/필드','display값','typed값','formula']);
    i60AppendStructure_(detail, prod);
    i60AppendStructure_(detail, backup);
    i60AppendStructure_(detail, preview);
    i60AppendCompare_(detail, 'CURRENT_vs_BACKUP', cb);
    i60AppendCompare_(detail, 'CURRENT_vs_PREVIEW', cp);
    var dsh=i60Ensure_(ss,I60_DETAIL);
    dsh.clearContents();
    if (detail.length) dsh.getRange(1,1,detail.length,8).setValues(detail);
    dsh.setFrozenRows(1);
    dsh.setColumnWidth(1,180);dsh.setColumnWidth(2,240);dsh.setColumnWidth(3,70);dsh.setColumnWidth(4,70);
    dsh.setColumnWidth(5,230);dsh.setColumnWidth(6,420);dsh.setColumnWidth(7,420);dsh.setColumnWidth(8,320);

    names.forEach(function(name){
      var after=i60SheetSigDisplay_(src[name]);
      if (after!==before[name]) throw new Error('보호시트 display signature 변경: '+name);
    });

    i60WriteStatus_(statusSh,'PASS','DONE','Issue59 rollback 구조/논리 상태 read-only 진단 완료',{
      verdict:verdict,
      prod:prod,backup:backup,preview:preview,
      cb:cb,cp:cp,
      oldStats:oldStats?'YES':'NO',newStats:newStats?'YES':'NO',
      backupOld:backupOld?'YES':'NO',previewNew:previewNew?'YES':'NO'
    });
    return {ok:true,done:true,verdict:verdict};
  } catch(e) {
    var msg=String(e&&e.message?e.message:e);
    try { i60WriteStatus_(statusSh,'ERROR','FAILED','Issue60 v1.1 read-only 구조진단 실패',{error:msg}); } catch(ignore) {}
    throw e;
  }
}

function runLotteonRemoteTaskContinueRemote_(){ return runLotteonRemoteTaskStartRemote_(); }

function i60Inspect_(sheet){
  var rows=sheet.getLastRow(), cols=sheet.getLastColumn();
  var vals=rows&&cols?sheet.getRange(1,1,rows,cols).getValues():[];
  var disp=rows&&cols?sheet.getRange(1,1,rows,cols).getDisplayValues():[];
  var formulas=rows&&cols?sheet.getRange(1,1,rows,cols).getFormulas():[];
  var header=i60FindHeaderRow_(disp);
  var stats=i60StatsFlexible_(sheet,disp,header);
  return {name:sheet.getName(),rows:rows,cols:cols,values:vals,display:disp,formulas:formulas,header:header,stats:stats};
}

function i60FindHeaderRow_(disp){
  var best={row:0,score:-1,headers:[],account:-1,order:-1,status:-1,purchase:-1,v669:-1,v670:-1,canonical:-1};
  var max=Math.min(disp.length,20);
  for(var r=0;r<max;r++){
    var h=(disp[r]||[]).map(i60Text_);
    var idx={
      account:i60Find_(h,['쿠팡계정ID']),
      order:i60Find_(h,['주문번호','마켓주문번호']),
      status:i60Find_(h,['카드매칭상태']),
      purchase:i60Find_(h,['주문매입금액','매입금액']),
      v669:i60Find_(h,['v6.69 2차귀속']),
      v670:i60Find_(h,['v6.70 3차귀속']),
      canonical:i60Find_(h,['canonicalEvidenceKey'])
    };
    var score=0;
    if(idx.account>=0)score+=3;if(idx.order>=0)score+=3;if(idx.status>=0)score+=2;if(idx.purchase>=0)score+=2;
    if(idx.v669>=0)score++;if(idx.v670>=0)score++;if(idx.canonical>=0)score++;
    if(score>best.score){best={row:r+1,score:score,headers:h,account:idx.account,order:idx.order,status:idx.status,purchase:idx.purchase,v669:idx.v669,v670:idx.v670,canonical:idx.canonical};}
  }
  return best;
}

function i60StatsFlexible_(sheet,disp,header){
  var st={dataRows:0,keyedRows:0,blankRows:0,partialKeyRows:0,matched:0,nonCard:0,ambiguous:0,noMatch:0,v669:0,v670:0,purchase:0,duplicateKeys:0,keys:{},firstDataRow:0,firstBlankKeyRow:0};
  if(!header || header.row<1 || header.account<0 || header.order<0) return st;
  var seen={};
  for(var r=header.row;r<disp.length;r++){
    var row=disp[r]||[];
    var nonempty=false;for(var c=0;c<row.length;c++){if(i60Text_(row[c])){nonempty=true;break;}}
    if(!nonempty){st.blankRows++;continue;}
    st.dataRows++;if(!st.firstDataRow)st.firstDataRow=r+1;
    var a=i60Text_(row[header.account]), o=i60Text_(row[header.order]);
    if(!a || !o){st.partialKeyRows++;if(!st.firstBlankKeyRow)st.firstBlankKeyRow=r+1;continue;}
    var k=i60Key_(a,o);st.keyedRows++;if(seen[k])st.duplicateKeys++;seen[k]=true;st.keys[k]=i60Material_(row,header);
    if(header.status>=0){var s=i60Text_(row[header.status]);if(s==='MATCHED'||s==='MASTER_MATCHED')st.matched++;else if(s==='NON_CARD')st.nonCard++;else if(s==='AMBIGUOUS')st.ambiguous++;else st.noMatch++;}
    if(header.v669>=0&&i60Text_(row[header.v669])==='Y')st.v669++;
    if(header.v670>=0&&i60Text_(row[header.v670])==='Y')st.v670++;
    if(header.purchase>=0)st.purchase+=i60Num_(row[header.purchase]);
  }
  return st;
}

function i60Material_(row,h){
  var fields=['account','order','status','purchase','v669','v670','canonical'];
  var out={};fields.forEach(function(f){var ix=h[f];out[f]=ix>=0?i60Text_(row[ix]):'';});
  return out;
}

function i60CompareByKey_(left,right){
  var a=left.stats.keys||{}, b=right.stats.keys||{};var out={overlap:0,onlyLeft:0,onlyRight:0,materialDiff:0,samples:[]};
  Object.keys(a).forEach(function(k){if(b[k]){out.overlap++;if(i60MaterialSig_(a[k])!==i60MaterialSig_(b[k])){out.materialDiff++;if(out.samples.length<20)out.samples.push({key:k,left:a[k],right:b[k]});}}else out.onlyLeft++;});
  Object.keys(b).forEach(function(k){if(!a[k])out.onlyRight++;});return out;
}
function i60MaterialSig_(x){return [x.status,Math.round(i60Num_(x.purchase)),x.v669,x.v670,x.canonical].join('|');}
function i60IsOldStats_(s){return s.keyedRows===1355&&s.matched===810&&s.nonCard===494&&s.ambiguous===1&&s.noMatch===50&&Math.round(s.purchase)===54807644;}
function i60IsNewStats_(s){return s.keyedRows===1355&&s.matched===808&&s.nonCard===498&&s.ambiguous===0&&s.noMatch===49&&s.v669===1161&&s.v670===81&&Math.round(s.purchase)===105762969;}

function i60AppendStructure_(detail,x){
  detail.push(['META',x.name,0,0,'size',x.rows+'x'+x.cols,'headerRow='+x.header.row+' score='+x.header.score,'']);
  detail.push(['META',x.name,0,0,'stats','keyed='+x.stats.keyedRows+' partial='+x.stats.partialKeyRows+' data='+x.stats.dataRows,'M='+x.stats.matched+' NC='+x.stats.nonCard+' A='+x.stats.ambiguous+' NM='+x.stats.noMatch+' purchase='+x.stats.purchase,'']);
  var maxR=Math.min(x.rows,12), maxC=Math.min(x.cols,30);
  for(var r=0;r<maxR;r++){
    var nonempty=false;for(var c=0;c<maxC;c++){if(i60Text_(x.display[r][c])){nonempty=true;break;}}
    if(!nonempty){detail.push(['ROW_SAMPLE',x.name,r+1,0,'(entire row blank)','','','']);continue;}
    for(var c2=0;c2<maxC;c2++){
      var dv=i60Text_(x.display[r][c2]), tv=i60Typed_(x.values[r][c2]), fm=i60Text_(x.formulas[r][c2]);
      if(dv||fm||r+1===x.header.row) detail.push(['CELL_SAMPLE',x.name,r+1,c2+1,(r+1===x.header.row?'HEADER':'')+(x.header.headers[c2]?' '+x.header.headers[c2]:''),dv,tv,fm]);
    }
  }
}
function i60AppendCompare_(detail,label,cmp){
  detail.push(['COMPARE',label,0,0,'summary','overlap='+cmp.overlap+' onlyLeft='+cmp.onlyLeft+' onlyRight='+cmp.onlyRight,'materialDiff='+cmp.materialDiff,'']);
  cmp.samples.forEach(function(s,i){detail.push(['COMPARE_SAMPLE',label,i+1,0,s.key,JSON.stringify(s.left),JSON.stringify(s.right),'']);});
}

function i60WriteStatus_(sheet,status,stage,message,x){
  x=x||{};var p=x.prod||{header:{},stats:{}},b=x.backup||{header:{},stats:{}},n=x.preview||{header:{},stats:{}},cb=x.cb||{},cp=x.cp||{};
  var rows=[['항목','값'],['버전',I60_VERSION],['상태',status],['단계',stage],['메시지',message],['판정',x.verdict||''],
    ['현재_lastRow',p.rows||0],['현재_lastCol',p.cols||0],['현재_headerRow',(p.header||{}).row||0],['현재_headerScore',(p.header||{}).score||0],['현재_firstDataRow',(p.stats||{}).firstDataRow||0],['현재_firstBlankKeyRow',(p.stats||{}).firstBlankKeyRow||0],['현재_dataRows',(p.stats||{}).dataRows||0],['현재_keyedRows',(p.stats||{}).keyedRows||0],['현재_partialKeyRows',(p.stats||{}).partialKeyRows||0],['현재_MATCHED',(p.stats||{}).matched||0],['현재_NON_CARD',(p.stats||{}).nonCard||0],['현재_AMBIGUOUS',(p.stats||{}).ambiguous||0],['현재_NO_MATCH',(p.stats||{}).noMatch||0],['현재_v6.69',(p.stats||{}).v669||0],['현재_v6.70',(p.stats||{}).v670||0],['현재매입합계',(p.stats||{}).purchase||0],['현재정규화중복',(p.stats||{}).duplicateKeys||0],
    ['백업_headerRow',(b.header||{}).row||0],['백업_keyedRows',(b.stats||{}).keyedRows||0],['백업_partialKeyRows',(b.stats||{}).partialKeyRows||0],['백업_MATCHED',(b.stats||{}).matched||0],['백업_NON_CARD',(b.stats||{}).nonCard||0],['백업_AMBIGUOUS',(b.stats||{}).ambiguous||0],['백업_NO_MATCH',(b.stats||{}).noMatch||0],['백업매입합계',(b.stats||{}).purchase||0],['백업_old집계일치',x.backupOld||''],
    ['preview_headerRow',(n.header||{}).row||0],['preview_keyedRows',(n.stats||{}).keyedRows||0],['preview_partialKeyRows',(n.stats||{}).partialKeyRows||0],['preview_MATCHED',(n.stats||{}).matched||0],['preview_NON_CARD',(n.stats||{}).nonCard||0],['preview_AMBIGUOUS',(n.stats||{}).ambiguous||0],['preview_NO_MATCH',(n.stats||{}).noMatch||0],['preview매입합계',(n.stats||{}).purchase||0],['preview_new집계일치',x.previewNew||''],
    ['현재/백업_overlap',cb.overlap||0],['현재/백업_material행차이',cb.materialDiff||0],['backupOnly',cb.onlyRight||0],['currentOnly_vs_backup',cb.onlyLeft||0],['현재/preview_overlap',cp.overlap||0],['현재/preview_material행차이',cp.materialDiff||0],['previewOnly',cp.onlyRight||0],['currentOnly_vs_preview',cp.onlyLeft||0],['old집계일치',x.oldStats||''],['new집계일치',x.newStats||''],
    ['부가세_카드매칭검증 변경','0'],['Issue59백업 변경','0'],['Issue54preview 변경','0'],['부가세_신고자료 변경','0'],['부가세_기간별 변경','0'],['카드사용내역_붙여넣기 변경','0'],['카드_마스터 변경','0'],['오류',x.error||''],['완료시각',(status==='PASS'||status==='ERROR')?new Date().toISOString():''],['갱신시각',new Date().toISOString()]];
  sheet.clearContents();sheet.getRange(1,1,rows.length,2).setValues(rows);sheet.setFrozenRows(1);sheet.setColumnWidth(1,250);sheet.setColumnWidth(2,650);
}

function i60SheetSigDisplay_(sheet){if(!sheet)return'MISSING';var r=sheet.getLastRow(),c=sheet.getLastColumn();if(!r||!c)return'EMPTY|'+r+'|'+c;var v=sheet.getRange(1,1,r,c).getDisplayValues();var h=2166136261;for(var i=0;i<v.length;i++){for(var j=0;j<v[i].length;j++){var s=i60Text_(v[i][j])+'\u001f';for(var k=0;k<s.length;k++){h^=s.charCodeAt(k);h=Math.imul(h,16777619);}}}return r+'x'+c+'|'+(h>>>0).toString(16);}
function i60Typed_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return'DATE:'+v.toISOString();return(typeof v).toUpperCase()+':'+i60Text_(v);}
function i60Need_(ss,name){var s=ss.getSheetByName(name);if(!s)throw new Error('필수 시트 없음: '+name);return s;}
function i60Ensure_(ss,name){return ss.getSheetByName(name)||ss.insertSheet(name);}
function i60Find_(h,names){for(var i=0;i<names.length;i++){var x=h.indexOf(names[i]);if(x>=0)return x;}return-1;}
function i60Text_(v){return String(v==null?'':v).trim();}
function i60Num_(v){var n=Number(i60Text_(v).replace(/[,원\s]/g,''));return isNaN(n)?0:n;}
function i60Key_(a,o){var aa=i60Text_(a).toLowerCase();var oo=i60Text_(o).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');return aa&&oo?aa+'|'+oo:'';}
