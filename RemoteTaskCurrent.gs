/**
 * Issue #60 v1.0 read-only recovery diagnostic after Issue59 ROLLBACK_ERROR.
 * Writes only ISSUE60 diagnostic/status sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE60-v1.0-20260813',
  title: 'Issue59 ROLLBACK_ERROR 운영 카드검증 복구상태 read-only 진단',
  enabled: true,
  statusSheet: 'ISSUE60_복구진단상태'
};

var I60_VERSION = 'v1.0-ISSUE60-ROLLBACK-STATE-DIAGNOSTIC';
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
  i60WriteStatus_(statusSh, 'RUNNING', 'CHECK', 'Issue59 rollback 상태 read-only 진단 시작', {});

  var prod = i60Need_(ss, I60_PROD);
  var backup = i60Need_(ss, I60_BACKUP);
  var preview = i60Need_(ss, I60_PREVIEW);
  var vat = i60Need_(ss, I60_VAT);
  var period = i60Need_(ss, I60_PERIOD);
  var history = i60Need_(ss, I60_HISTORY);
  var master = i60Need_(ss, I60_MASTER);

  var before = {
    prod:i60SheetSig_(prod), backup:i60SheetSig_(backup), preview:i60SheetSig_(preview),
    vat:i60SheetSig_(vat), period:i60SheetSig_(period), history:i60SheetSig_(history), master:i60SheetSig_(master)
  };

  try {
    var oldStats = i60DiagStats_(backup);
    i60AssertStats_('Issue59 backup', oldStats, {
      orders:1355, matched:810, nonCard:494, ambiguous:1, noMatch:50, purchase:54807644
    });
    var newStats = i60DiagStats_(preview);
    i60AssertStats_('Issue54 preview', newStats, {
      orders:1355, matched:808, nonCard:498, ambiguous:0, noMatch:49, v669:1161, v670:81, purchase:105762969
    });
    var curStats = i60DiagStats_(prod);

    var strictOld = i60CompareCells_(backup, prod, 'BACKUP_vs_CURRENT', 40);
    var strictNew = i60CompareCells_(preview, prod, 'PREVIEW_vs_CURRENT', 40);
    var logicalOld = i60CompareLogical_(backup, prod, 'BACKUP_vs_CURRENT', 40);
    var logicalNew = i60CompareLogical_(preview, prod, 'PREVIEW_vs_CURRENT', 40);

    var oldStatsMatch = i60StatsMatch_(curStats, {
      orders:1355, matched:810, nonCard:494, ambiguous:1, noMatch:50, purchase:54807644
    });
    var newStatsMatch = i60StatsMatch_(curStats, {
      orders:1355, matched:808, nonCard:498, ambiguous:0, noMatch:49, v669:1161, v670:81, purchase:105762969
    });

    var classification = 'MIXED_OR_UNKNOWN';
    if (logicalOld.mismatchRows === 0 && logicalOld.oldOnly === 0 && logicalOld.newOnly === 0) {
      classification = 'LOGICALLY_OLD_ROLLBACK_OK';
    } else if (logicalNew.mismatchRows === 0 && logicalNew.oldOnly === 0 && logicalNew.newOnly === 0) {
      classification = 'LOGICALLY_NEW_APPLY_PRESENT';
    } else if (oldStatsMatch) {
      classification = 'OLD_STATS_BUT_CELL_DIFF';
    } else if (newStatsMatch) {
      classification = 'NEW_STATS_BUT_CELL_DIFF';
    }

    var details = [];
    details.push(['구분','비교','행','열','헤더','왼쪽 strict','오른쪽 strict','왼쪽 표시','오른쪽 표시','주문키','비고']);
    strictOld.samples.forEach(function(x){ details.push(i60DetailRow_(x)); });
    strictNew.samples.forEach(function(x){ details.push(i60DetailRow_(x)); });
    logicalOld.samples.forEach(function(x){ details.push(i60LogicalDetailRow_(x)); });
    logicalNew.samples.forEach(function(x){ details.push(i60LogicalDetailRow_(x)); });
    i60WriteDetail_(ss, details);

    var after = {
      prod:i60SheetSig_(prod), backup:i60SheetSig_(backup), preview:i60SheetSig_(preview),
      vat:i60SheetSig_(vat), period:i60SheetSig_(period), history:i60SheetSig_(history), master:i60SheetSig_(master)
    };
    Object.keys(before).forEach(function(k){
      if (before[k] !== after[k]) throw new Error('보호시트 signature 변경: ' + k);
    });

    i60WriteStatus_(statusSh, 'PASS', 'DONE', 'Issue59 rollback 상태 read-only 진단 완료', {
      classification:classification,
      currentOrders:curStats.orders,
      currentMatched:curStats.matched,
      currentNonCard:curStats.nonCard,
      currentAmbiguous:curStats.ambiguous,
      currentNoMatch:curStats.noMatch,
      currentV669:curStats.v669,
      currentV670:curStats.v670,
      currentPurchase:curStats.purchase,
      currentDup:curStats.duplicateKeys,
      overlapBackup:i60Overlap_(curStats.keys, oldStats.keys),
      overlapPreview:i60Overlap_(curStats.keys, newStats.keys),
      strictOld:strictOld.strictMismatch,
      displayOld:strictOld.displayMismatch,
      strictNew:strictNew.strictMismatch,
      displayNew:strictNew.displayMismatch,
      logicalOld:logicalOld.mismatchRows,
      logicalNew:logicalNew.mismatchRows,
      oldOnlyBackup:logicalOld.oldOnly,
      currentOnlyBackup:logicalOld.newOnly,
      oldOnlyPreview:logicalNew.oldOnly,
      currentOnlyPreview:logicalNew.newOnly,
      oldStatsMatch:oldStatsMatch ? 'YES':'NO',
      newStatsMatch:newStatsMatch ? 'YES':'NO',
      prodChange:'0', backupChange:'0', previewChange:'0', vatChange:'0', periodChange:'0', historyChange:'0', masterChange:'0'
    });
    return {ok:true, done:true, classification:classification};
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    i60WriteStatus_(statusSh, 'ERROR', 'FAILED', 'Issue60 read-only 진단 실패', {error:msg});
    throw e;
  }
}

function runLotteonRemoteTaskContinueRemote_() {
  return runLotteonRemoteTaskStartRemote_();
}

function i60Need_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 1) throw new Error('필수 시트 없음: ' + name);
  return sh;
}

function i60DiagStats_(sheet) {
  var v = sheet.getDataRange().getValues();
  if (v.length < 2) throw new Error(sheet.getName() + ' 데이터가 없습니다.');
  var h = v[0].map(i60Text_);
  var ix = {
    account:i60Find_(h,['쿠팡계정ID','계정ID','마켓아이디']),
    order:i60Find_(h,['주문번호','마켓주문번호']),
    purchase:i60Find_(h,['주문매입금액','매입금액']),
    status:i60Find_(h,['카드매칭상태','매칭상태']),
    v669:i60Find_(h,['v6.69 2차귀속']),
    v670:i60Find_(h,['v6.70 3차귀속'])
  };
  if (ix.account<0 || ix.order<0 || ix.purchase<0 || ix.status<0) throw new Error(sheet.getName() + ' 필수 헤더 누락');
  var st={orders:0,matched:0,nonCard:0,ambiguous:0,noMatch:0,v669:0,v670:0,purchase:0,duplicateKeys:0,keys:{}};
  for (var r=1;r<v.length;r++) {
    var key=i60Key_(v[r][ix.account],v[r][ix.order]);
    if (!key) throw new Error(sheet.getName() + ' 주문키 공란 R' + (r+1));
    if (st.keys[key]) st.duplicateKeys++;
    st.keys[key]=true;
    st.orders++;
    var s=i60Status_(v[r][ix.status]);
    if (s==='MATCHED') st.matched++;
    else if (s==='NON_CARD') st.nonCard++;
    else if (s==='AMBIGUOUS') st.ambiguous++;
    else st.noMatch++;
    if (ix.v669>=0 && i60Text_(v[r][ix.v669])==='Y') st.v669++;
    if (ix.v670>=0 && i60Text_(v[r][ix.v670])==='Y') st.v670++;
    st.purchase += i60Num_(v[r][ix.purchase]);
  }
  return st;
}

function i60AssertStats_(label, a, e) {
  Object.keys(e).forEach(function(k){
    if (Math.round(Number(a[k]||0)) !== Math.round(Number(e[k]||0))) {
      throw new Error(label + ' ' + k + ' 불일치: ' + a[k] + ' / 기대 ' + e[k]);
    }
  });
}
function i60StatsMatch_(a,e) {
  return Object.keys(e).every(function(k){return Math.round(Number(a[k]||0))===Math.round(Number(e[k]||0));});
}

function i60CompareCells_(left, right, label, sampleLimit) {
  var lr=left.getLastRow(), lc=left.getLastColumn(), rr=right.getLastRow(), rc=right.getLastColumn();
  var rows=Math.max(lr,rr), cols=Math.max(lc,rc);
  var lv=lr&&lc?left.getRange(1,1,lr,lc).getValues():[];
  var rv=rr&&rc?right.getRange(1,1,rr,rc).getValues():[];
  var ld=lr&&lc?left.getRange(1,1,lr,lc).getDisplayValues():[];
  var rd=rr&&rc?right.getRange(1,1,rr,rc).getDisplayValues():[];
  var headers=[];
  for(var c=0;c<cols;c++) headers[c]=i60Text_((lv[0]||[])[c] || (rv[0]||[])[c]);
  var strict=0, display=0, samples=[];
  for(var r=0;r<rows;r++) for(var c=0;c<cols;c++) {
    var l = (r<lr && c<lc) ? lv[r][c] : '__MISSING_CELL__';
    var q = (r<rr && c<rc) ? rv[r][c] : '__MISSING_CELL__';
    var ls = (r<lr && c<lc) ? ld[r][c] : '__MISSING_CELL__';
    var rs = (r<rr && c<rc) ? rd[r][c] : '__MISSING_CELL__';
    var stDiff = i60Strict_(l)!==i60Strict_(q);
    var dsDiff = String(ls)!==String(rs);
    if (stDiff) strict++;
    if (dsDiff) display++;
    if ((stDiff || dsDiff) && samples.length<sampleLimit) {
      samples.push({kind:'CELL',label:label,row:r+1,col:c+1,header:headers[c]||'',leftStrict:i60Strict_(l),rightStrict:i60Strict_(q),leftDisplay:String(ls),rightDisplay:String(rs),key:'',note:(stDiff?'STRICT ':'')+(dsDiff?'DISPLAY':'')});
    }
  }
  return {strictMismatch:strict,displayMismatch:display,samples:samples,leftRows:lr,leftCols:lc,rightRows:rr,rightCols:rc};
}

function i60CompareLogical_(left, right, label, sampleLimit) {
  var a=i60LogicalMap_(left), b=i60LogicalMap_(right);
  var mismatch=0, oldOnly=0, newOnly=0, samples=[];
  Object.keys(a.map).forEach(function(k){
    if (!b.map[k]) {
      oldOnly++;
      if(samples.length<sampleLimit)samples.push({kind:'LOGICAL',label:label,key:k,header:'',leftStrict:'ROW_PRESENT',rightStrict:'ROW_MISSING',leftDisplay:'',rightDisplay:'',note:'LEFT_ONLY'});
      return;
    }
    if (a.map[k].fingerprint!==b.map[k].fingerprint) {
      mismatch++;
      if(samples.length<sampleLimit){
        var d=i60FirstFieldDiff_(a.map[k],b.map[k]);
        samples.push({kind:'LOGICAL',label:label,key:k,header:d.header,leftStrict:d.left,rightStrict:d.right,leftDisplay:'',rightDisplay:'',note:'MATERIAL_FIELD_DIFF'});
      }
    }
  });
  Object.keys(b.map).forEach(function(k){
    if(!a.map[k]){
      newOnly++;
      if(samples.length<sampleLimit)samples.push({kind:'LOGICAL',label:label,key:k,header:'',leftStrict:'ROW_MISSING',rightStrict:'ROW_PRESENT',leftDisplay:'',rightDisplay:'',note:'RIGHT_ONLY'});
    }
  });
  return {mismatchRows:mismatch,oldOnly:oldOnly,newOnly:newOnly,samples:samples};
}

function i60LogicalMap_(sheet) {
  var v=sheet.getDataRange().getValues();
  var h=v[0].map(i60Text_);
  var account=i60Find_(h,['쿠팡계정ID','계정ID','마켓아이디']);
  var order=i60Find_(h,['주문번호','마켓주문번호']);
  if(account<0||order<0) throw new Error(sheet.getName()+' 주문키 헤더 누락');
  var map={};
  for(var r=1;r<v.length;r++){
    var key=i60Key_(v[r][account],v[r][order]);
    if(map[key]) throw new Error(sheet.getName()+' 정규화 주문키 중복: '+key);
    var fields={}, parts=[];
    for(var c=0;c<h.length;c++){
      var header=h[c] || ('COL'+(c+1));
      if(c===account||c===order) continue;
      var sem=i60Semantic_(v[r][c],header);
      fields[header]=sem;
      parts.push(header+'='+sem);
    }
    map[key]={fields:fields,fingerprint:parts.join('\u001e')};
  }
  return {map:map};
}

function i60FirstFieldDiff_(a,b){
  var keys={}; Object.keys(a.fields).forEach(function(k){keys[k]=1;}); Object.keys(b.fields).forEach(function(k){keys[k]=1;});
  var names=Object.keys(keys).sort();
  for(var i=0;i<names.length;i++){
    var k=names[i], x=a.fields[k], y=b.fields[k];
    if(x!==y) return {header:k,left:x,right:y};
  }
  return {header:'',left:a.fingerprint,right:b.fingerprint};
}

function i60Semantic_(v, header) {
  var h=i60Text_(header);
  if (/금액|후보수|주문건수|매입금액|승인금액/.test(h)) return 'NUM:'+String(i60Num_(v));
  if (/일$|일자|날짜|승인일|주문일/.test(h)) return 'DATE:'+i60DateKey_(v);
  if (h==='카드매칭상태' || h==='매칭상태') return 'STATUS:'+i60Status_(v);
  return 'TEXT:'+i60Text_(v);
}

function i60DetailRow_(x){return [x.kind,x.label,x.row||'',x.col||'',x.header||'',x.leftStrict||'',x.rightStrict||'',x.leftDisplay||'',x.rightDisplay||'',x.key||'',x.note||''];}
function i60LogicalDetailRow_(x){return [x.kind,x.label,'','',x.header||'',x.leftStrict||'',x.rightStrict||'',x.leftDisplay||'',x.rightDisplay||'',x.key||'',x.note||''];}

function i60WriteDetail_(ss, rows) {
  var sh=ss.getSheetByName(I60_DETAIL)||ss.insertSheet(I60_DETAIL);
  sh.clearContents();
  if(rows.length) sh.getRange(1,1,rows.length,rows[0].length).setValues(rows);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,rows[0].length).setFontWeight('bold');
  sh.setColumnWidth(1,100); sh.setColumnWidth(2,180); sh.setColumnWidth(5,180); sh.setColumnWidth(6,260); sh.setColumnWidth(7,260); sh.setColumnWidth(10,220); sh.setColumnWidth(11,180);
}

function i60WriteStatus_(sh,status,stage,message,x){
  x=x||{};
  var rows=[
    ['항목','값'],['버전',I60_VERSION],['상태',status],['단계',stage],['메시지',message],
    ['판정',x.classification||''],
    ['현재운영주문',x.currentOrders||0],['현재_MATCHED',x.currentMatched||0],['현재_NON_CARD',x.currentNonCard||0],['현재_AMBIGUOUS',x.currentAmbiguous||0],['현재_NO_MATCH',x.currentNoMatch||0],
    ['현재_v6.69',x.currentV669||0],['현재_v6.70',x.currentV670||0],['현재매입합계',x.currentPurchase||0],['현재정규화중복',x.currentDup||0],
    ['현재/백업_overlap',x.overlapBackup||0],['현재/preview_overlap',x.overlapPreview||0],
    ['current_vs_backup_strict셀차이',x.strictOld||0],['current_vs_backup_display셀차이',x.displayOld||0],['current_vs_backup_material행차이',x.logicalOld||0],
    ['backupOnly',x.oldOnlyBackup||0],['currentOnly_vs_backup',x.currentOnlyBackup||0],
    ['current_vs_preview_strict셀차이',x.strictNew||0],['current_vs_preview_display셀차이',x.displayNew||0],['current_vs_preview_material행차이',x.logicalNew||0],
    ['previewOnly',x.oldOnlyPreview||0],['currentOnly_vs_preview',x.currentOnlyPreview||0],
    ['old집계일치',x.oldStatsMatch||''],['new집계일치',x.newStatsMatch||''],
    ['부가세_카드매칭검증 변경',x.prodChange||'0'],['Issue59백업 변경',x.backupChange||'0'],['Issue54preview 변경',x.previewChange||'0'],
    ['부가세_신고자료 변경',x.vatChange||'0'],['부가세_기간별 변경',x.periodChange||'0'],['카드사용내역_붙여넣기 변경',x.historyChange||'0'],['카드_마스터 변경',x.masterChange||'0'],
    ['오류',x.error||''],['완료시각',(status==='PASS'||status==='ERROR')?new Date().toISOString():''],['갱신시각',new Date().toISOString()]
  ];
  sh.clearContents(); sh.getRange(1,1,rows.length,2).setValues(rows); sh.setFrozenRows(1); sh.getRange(1,1,1,2).setFontWeight('bold'); sh.setColumnWidth(1,300); sh.setColumnWidth(2,700);
}

function i60SheetSig_(sheet){
  var rows=sheet.getLastRow(),cols=sheet.getLastColumn();
  if(!rows||!cols)return 'EMPTY|'+rows+'|'+cols;
  var vals=sheet.getRange(1,1,rows,cols).getValues(),h=2166136261;
  for(var r=0;r<vals.length;r++)for(var c=0;c<vals[r].length;c++){
    var s=i60Strict_(vals[r][c])+'\u001f';
    for(var i=0;i<s.length;i++){h^=s.charCodeAt(i);h=Math.imul(h,16777619);}
  }
  return rows+'x'+cols+'|'+(h>>>0).toString(16);
}
function i60Strict_(v){
  if(v==='__MISSING_CELL__')return 'MISSING';
  if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return 'D:'+v.toISOString();
  if(typeof v==='number')return 'N:'+String(v);
  if(typeof v==='boolean')return 'B:'+String(v);
  return 'T:'+i60Text_(v);
}
function i60Status_(v){var s=i60Text_(v);if(s==='MASTER_MATCHED')return 'MATCHED';if(s==='MATCHED'||s==='NON_CARD'||s==='AMBIGUOUS'||s==='NO_MATCH')return s;return s||'NO_MATCH';}
function i60DateKey_(v){
  if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd HH:mm:ss');
  var s=i60Text_(v); var m=s.match(/(20\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if(!m)return s;
  return m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2)+(m[4]?' '+('0'+m[4]).slice(-2)+':'+m[5]+':'+(m[6]||'00'):'');
}
function i60Key_(account,order){var a=i60Text_(account).toLowerCase();var o=i60Text_(order).replace(/[\s._()\[\]{}\-\/]/g,'').replace(/^0+(?=\d)/,'');return a&&o?a+'|'+o:'';}
function i60Num_(v){var n=Number(typeof v==='number'?v:i60Text_(v).replace(/[,원\s]/g,''));return isNaN(n)?0:n;}
function i60Text_(v){return String(v==null?'':v).trim();}
function i60Find_(h,names){for(var i=0;i<names.length;i++){var x=h.indexOf(names[i]);if(x>=0)return x;}return -1;}
function i60Overlap_(a,b){var n=0;Object.keys(a||{}).forEach(function(k){if(b&&b[k])n++;});return n;}
function i60Ensure_(ss,name){return ss.getSheetByName(name)||ss.insertSheet(name);}
