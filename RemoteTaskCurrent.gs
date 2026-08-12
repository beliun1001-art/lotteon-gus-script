/**
 * Issue #48 v1.1 standalone diagnostic.
 * Audits all 2026 H1 orders between current VAT detail and card verification
 * using normalized order-number JOIN keys.
 * IMPORTANT: header lookup and number parsing mirror production v6.60 semantics.
 * Reads production sheets; writes only ISSUE48_* diagnostic sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE48-v1.1-20260812',
  title: '상반기 VAT 현재 상세 ↔ 카드매칭검증 1,355건 stale audit v1.1',
  enabled: true,
  inputSheet: 'ISSUE47_3단계JOIN진단',
  outputSheet: 'ISSUE48_VAT검증STALE감사',
  statusSheet: 'ISSUE48_진단상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var state = issue48v11Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue48v11Write_(state,[
    ['항목','값'],
    ['버전','v1.1-ISSUE48-H1-VAT-VERIFY-STALE-AUDIT'],
    ['상태','RUNNING'],['단계','LOAD'],
    ['메시지','production-compatible stale audit 시작'],
    ['운영시트 변경','0'],['갱신시각',new Date().toISOString()]
  ]);

  try {
    var vat = ss.getSheetByName('부가세_신고자료');
    var verify = ss.getSheetByName('부가세_카드매칭검증');
    var issue47 = ss.getSheetByName(LOTTEON_REMOTE_TASK.inputSheet);
    if (!vat || vat.getLastRow() < 2) throw new Error('부가세_신고자료 시트가 없습니다.');
    if (!verify || verify.getLastRow() < 2) throw new Error('부가세_카드매칭검증 시트가 없습니다.');
    if (!issue47 || issue47.getLastRow() < 2) throw new Error('ISSUE47_3단계JOIN진단 시트가 없습니다.');

    var targetSet = issue48v11TargetSet_(issue47);
    if (Object.keys(targetSet).length !== 5) throw new Error('Issue47 대상 건수 불일치: 기대 5건, 실제 ' + Object.keys(targetSet).length + '건');

    var vv = vat.getDataRange().getValues();
    var vh = vv[0] || [];
    var vi = {
      year:issue48v11FindFirst_(vh,['신고연도']),
      half:issue48v11FindFirst_(vh,['반기']),
      business:issue48v11FindFirst_(vh,['사업자등록번호']),
      account:issue48v11FindFirst_(vh,['쿠팡계정ID','마켓아이디','계정ID']),
      order:issue48v11FindFirst_(vh,['주문번호','마켓주문번호','주문ID','주문ID(마켓)']),
      date:issue48v11FindFirst_(vh,['날짜','주문일','주문일자','마켓주문일자']),
      payment:issue48v11FindFirst_(vh,['롯데결제수단','구매결제수단','결제수단']),
      purchase:issue48v11FindFirst_(vh,['매입금액'])
    };
    ['year','half','business','account','order','purchase'].forEach(function(k){if(vi[k]<0) throw new Error('VAT 필수 헤더 누락: '+k);});

    var cv = verify.getDataRange().getValues();
    var ch = cv[0] || [];
    var ci = {
      year:issue48v11FindFirst_(ch,['신고연도']),
      half:issue48v11FindFirst_(ch,['반기']),
      business:issue48v11FindFirst_(ch,['사업자등록번호']),
      account:issue48v11FindFirst_(ch,['쿠팡계정ID','마켓아이디','계정ID']),
      order:issue48v11FindFirst_(ch,['주문번호','마켓주문번호','주문ID','주문ID(마켓)']),
      date:issue48v11FindFirst_(ch,['주문일','날짜','주문일자']),
      payment:issue48v11FindFirst_(ch,['롯데결제수단','구매결제수단','결제수단']),
      purchase:issue48v11FindFirst_(ch,['주문매입금액']),
      status:issue48v11FindFirst_(ch,['카드매칭상태'])
    };
    ['year','half','business','account','order','purchase','status'].forEach(function(k){if(ci[k]<0) throw new Error('검증 필수 헤더 누락: '+k);});

    var vatPurchaseCandidates = issue48v11HeaderCandidates_(vh,['매입금액']);
    var verifyPurchaseCandidates = issue48v11HeaderCandidates_(ch,['주문매입금액']);

    var vatOrders = issue48v11GroupVat_(vv, vi);
    var verifyOrders = issue48v11ReadVerify_(cv, ci);
    var vatMap = issue48v11Bucket_(vatOrders);
    var verifyMap = issue48v11Bucket_(verifyOrders);
    var allKeys = {};
    Object.keys(vatMap).forEach(function(k){allKeys[k]=true;});
    Object.keys(verifyMap).forEach(function(k){allKeys[k]=true;});

    var headers = [
      'JOIN키','사업자등록번호','쿠팡계정ID','정규화주문번호','VAT주문번호','검증주문번호',
      'VAT주문그룹수','검증행수','VAT상세행수','VAT매입금액','검증주문매입금액','차액(VAT-검증)',
      '카드매칭상태','분류','Issue47대상','VAT문장부호','검증문장부호','메모'
    ];
    var out=[];
    var stats={
      vatOrders:vatOrders.length,verifyOrders:verifyOrders.length,oneToOne:0,vatOnly:0,verifyOnly:0,duplicateKeys:0,
      amountMatch:0,amountMismatch:0,vatPosVerifyZero:0,vatZeroVerifyPos:0,otherDiff:0,
      vatPurchase:0,verifyPurchase:0,targetCovered:0,targetVatPosVerifyZero:0
    };
    var statusCounts={}, mismatchStatusCounts={}, targetSeen={};

    vatOrders.forEach(function(x){stats.vatPurchase += x.purchase;});
    verifyOrders.forEach(function(x){
      stats.verifyPurchase += x.purchase;
      var s=x.status||'공란'; statusCounts[s]=(statusCounts[s]||0)+1;
    });

    Object.keys(allKeys).sort().forEach(function(key){
      var va=vatMap[key]||[], ca=verifyMap[key]||[];
      var vp=0,cp=0,vatDetailRows=0,status='',vatRaw='',verifyRaw='',classification='';
      va.forEach(function(x){vp+=x.purchase;vatDetailRows+=x.detailRows;});
      ca.forEach(function(x){cp+=x.purchase;});
      if(va.length) vatRaw=issue48v11Distinct_(va.map(function(x){return x.orderRaw;})).join(' | ');
      if(ca.length) verifyRaw=issue48v11Distinct_(ca.map(function(x){return x.orderRaw;})).join(' | ');
      if(ca.length) status=issue48v11Distinct_(ca.map(function(x){return x.status;})).join(' | ');

      if(!va.length){classification='검증_ONLY';stats.verifyOnly++;}
      else if(!ca.length){classification='VAT_ONLY';stats.vatOnly++;}
      else if(va.length!==1||ca.length!==1){classification='정규화키중복';stats.duplicateKeys++;}
      else{
        stats.oneToOne++;
        if(Math.round(vp)===Math.round(cp)){classification='매입금액일치';stats.amountMatch++;}
        else{
          stats.amountMismatch++;
          if(vp>0&&cp===0){classification='VAT양수_검증0';stats.vatPosVerifyZero++;}
          else if(vp===0&&cp>0){classification='VAT0_검증양수';stats.vatZeroVerifyPos++;}
          else{classification='기타금액차이';stats.otherDiff++;}
          var ms=status||'공란';mismatchStatusCounts[ms]=(mismatchStatusCounts[ms]||0)+1;
        }
      }

      var sample=va.length?va[0]:ca[0];
      var tkey=issue48v11TargetKey_(sample.account,sample.orderNorm);
      var isTarget=!!targetSet[tkey];
      if(isTarget&&!targetSeen[tkey]){
        targetSeen[tkey]=true;stats.targetCovered++;
        if(classification==='VAT양수_검증0')stats.targetVatPosVerifyZero++;
      }

      if(classification!=='매입금액일치'){
        var memo=[];
        if(va.length!==1)memo.push('VAT주문그룹='+va.length);
        if(ca.length!==1)memo.push('검증행='+ca.length);
        if(vatRaw&&verifyRaw&&vatRaw!==verifyRaw&&issue48v11OrderNorm_(vatRaw)===issue48v11OrderNorm_(verifyRaw))memo.push('주문번호표현만다름');
        out.push([
          key,sample.business,sample.account,sample.orderNorm,vatRaw,verifyRaw,va.length,ca.length,vatDetailRows,
          vp,cp,vp-cp,status,classification,isTarget?'Y':'N',issue48v11Punct_(vatRaw),issue48v11Punct_(verifyRaw),memo.join(' / ')
        ]);
      }
    });

    var output=issue48v11Ensure_(ss,LOTTEON_REMOTE_TASK.outputSheet);
    output.clearContents();
    output.getRange(1,1,1,headers.length).setValues([headers]);
    if(out.length)output.getRange(2,1,out.length,headers.length).setValues(out);
    output.setFrozenRows(1);
    output.getRange(1,1,1,headers.length).setBackground('#d9eaf7').setFontWeight('bold');
    if(out.length)output.getRange(2,10,out.length,3).setNumberFormat('#,##0');

    var statusRows=[
      ['항목','값'],
      ['버전','v1.1-ISSUE48-H1-VAT-VERIFY-STALE-AUDIT'],['상태','PASS'],['단계','DONE'],
      ['메시지','production-compatible stale audit 완료'],['운영시트 변경','0'],
      ['VAT매입금액선택열',issue48v11Col_(vi.purchase+1)+' / '+issue48v11Text_(vh[vi.purchase])],
      ['VAT매입금액헤더후보',vatPurchaseCandidates.join(' | ')||'없음'],
      ['검증매입금액선택열',issue48v11Col_(ci.purchase+1)+' / '+issue48v11Text_(ch[ci.purchase])],
      ['검증매입금액헤더후보',verifyPurchaseCandidates.join(' | ')||'없음'],
      ['VAT상반기주문',stats.vatOrders],['검증상반기주문',stats.verifyOrders],['정규화1대1매칭',stats.oneToOne],
      ['VAT_ONLY',stats.vatOnly],['검증_ONLY',stats.verifyOnly],['정규화키중복',stats.duplicateKeys],
      ['매입금액일치',stats.amountMatch],['매입금액불일치',stats.amountMismatch],
      ['VAT양수_검증0',stats.vatPosVerifyZero],['VAT0_검증양수',stats.vatZeroVerifyPos],['기타금액차이',stats.otherDiff],
      ['VAT총매입금액',Math.round(stats.vatPurchase)],['검증총매입금액',Math.round(stats.verifyPurchase)],
      ['총매입금액차액',Math.round(stats.vatPurchase-stats.verifyPurchase)],
      ['Issue47대상5_감사포함',stats.targetCovered],['Issue47대상5_VAT양수검증0',stats.targetVatPosVerifyZero],
      ['불일치출력건수',out.length]
    ];
    Object.keys(statusCounts).sort().forEach(function(k){statusRows.push(['현재상태_'+k,statusCounts[k]]);});
    Object.keys(mismatchStatusCounts).sort().forEach(function(k){statusRows.push(['금액불일치상태_'+k,mismatchStatusCounts[k]]);});
    statusRows.push(['완료시각',new Date().toISOString()]);
    issue48v11Write_(state,statusRows);
    try{MailApp.sendEmail('beliun1001@gmail.com','[LOTTEON 자동작업 결과][PASS] ISSUE48-v1.1',statusRows.map(function(x){return x[0]+': '+x[1];}).join('\n'));}catch(mailError){}
    return {ok:true,stats:stats,mismatches:out.length};
  }catch(e){
    issue48v11Write_(state,[
      ['항목','값'],['버전','v1.1-ISSUE48-H1-VAT-VERIFY-STALE-AUDIT'],['상태','ERROR'],['단계','FAILED'],
      ['메시지','production-compatible stale audit 실패'],['오류',String(e&&e.message?e.message:e)],
      ['운영시트 변경','0'],['갱신시각',new Date().toISOString()]
    ]);
    throw e;
  }
}

function issue48v11GroupVat_(values,ix){
  var exact={};
  for(var r=1;r<values.length;r++){
    var row=values[r],year=issue48v11Text_(row[ix.year]),half=issue48v11Text_(row[ix.half]);
    if(year!=='2026'||half!=='상반기')continue;
    var business=issue48v11Text_(row[ix.business]),account=issue48v11Text_(row[ix.account]),orderRaw=issue48v11Text_(row[ix.order]);
    var date=ix.date>=0?issue48v11Date_(row[ix.date]):'',payment=ix.payment>=0?issue48v11Text_(row[ix.payment]):'';
    var exactKey=orderRaw?[year,half,business,account,orderRaw].join('|'):[year,half,business,account,'BLANK',r].join('|');
    if(!exact[exactKey])exact[exactKey]={year:year,half:half,business:business,account:account,orderRaw:orderRaw,orderNorm:issue48v11OrderNorm_(orderRaw),date:date,payments:{},purchase:0,detailRows:0};
    exact[exactKey].purchase+=issue48v11Number_(row[ix.purchase]);
    exact[exactKey].detailRows++;
    if(payment)exact[exactKey].payments[payment]=true;
  }
  return Object.keys(exact).map(function(k){var x=exact[k];x.payment=Object.keys(x.payments).sort().join(', ');x.joinKey=issue48v11JoinKey_(x);return x;});
}
function issue48v11ReadVerify_(values,ix){
  var out=[];
  for(var r=1;r<values.length;r++){
    var row=values[r],year=issue48v11Text_(row[ix.year]),half=issue48v11Text_(row[ix.half]);
    if(year!=='2026'||half!=='상반기')continue;
    var x={year:year,half:half,business:issue48v11Text_(row[ix.business]),account:issue48v11Text_(row[ix.account]),orderRaw:issue48v11Text_(row[ix.order]),date:ix.date>=0?issue48v11Date_(row[ix.date]):'',payment:ix.payment>=0?issue48v11Text_(row[ix.payment]):'',purchase:issue48v11Number_(row[ix.purchase]),status:issue48v11Text_(row[ix.status]),detailRows:1};
    x.orderNorm=issue48v11OrderNorm_(x.orderRaw);x.joinKey=issue48v11JoinKey_(x);out.push(x);
  }
  return out;
}
function issue48v11Bucket_(rows){var m={};(rows||[]).forEach(function(x){if(!m[x.joinKey])m[x.joinKey]=[];m[x.joinKey].push(x);});return m;}
function issue48v11JoinKey_(x){
  var base=[x.year,x.half,x.business,x.account];
  if(x.orderNorm)return base.concat(['ORDER',x.orderNorm]).join('|');
  return base.concat(['BLANK',x.date,Math.round(x.purchase),issue48v11Compact_(x.payment)]).join('|');
}
function issue48v11TargetSet_(sheet){
  var v=sheet.getDataRange().getValues(),h=v[0]||[];
  var ni=issue48v11FindFirst_(h,['정규화주문번호']),ai=issue48v11FindFirst_(h,['쿠팡계정ID']);
  if(ni<0||ai<0)throw new Error('ISSUE47 대상 헤더 누락');
  var out={};for(var r=1;r<v.length;r++){var n=issue48v11Text_(v[r][ni]),a=issue48v11Text_(v[r][ai]);if(n)out[issue48v11TargetKey_(a,n)]=true;}return out;
}
function issue48v11TargetKey_(account,norm){return issue48v11Text_(account).toLowerCase()+'|'+issue48v11Text_(norm).toLowerCase();}
function issue48v11FindFirst_(headers,names){for(var n=0;n<names.length;n++){var wanted=issue48v11Compact_(names[n]);for(var i=0;i<headers.length;i++)if(issue48v11Compact_(headers[i])===wanted)return i;}return -1;}
function issue48v11HeaderCandidates_(headers,names){var wanted={};names.forEach(function(n){wanted[issue48v11Compact_(n)]=true;});var out=[];(headers||[]).forEach(function(h,i){if(wanted[issue48v11Compact_(h)])out.push(issue48v11Col_(i+1)+' / '+issue48v11Text_(h));});return out;}
function issue48v11Number_(v){if(typeof v==='number')return isNaN(v)?0:v;var n=Number(String(v==null?'0':v).replace(/[원,\s]/g,''));return isNaN(n)?0:n;}
function issue48v11Text_(v){return String(v==null?'':v).trim();}
function issue48v11Compact_(v){return issue48v11Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function issue48v11OrderNorm_(v){return issue48v11Text_(v).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');}
function issue48v11Distinct_(a){var m={},o=[];(a||[]).forEach(function(x){x=issue48v11Text_(x);if(!m[x]){m[x]=true;o.push(x);}});return o;}
function issue48v11Punct_(s){s=issue48v11Text_(s);var a=[];for(var i=0;i<s.length;i++){var c=s.charAt(i);if(!/[0-9a-zA-Z가-힣]/.test(c))a.push(c+'(U+'+('0000'+c.charCodeAt(0).toString(16).toUpperCase()).slice(-4)+')');}return a.join(' ');}
function issue48v11Date_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');return issue48v11Text_(v);}
function issue48v11Col_(n){var s='';while(n>0){var m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;}
function issue48v11Ensure_(ss,name){return ss.getSheetByName(name)||ss.insertSheet(name);}
function issue48v11Write_(sheet,rows){sheet.clearContents();sheet.getRange(1,1,rows.length,2).setValues(rows);sheet.setFrozenRows(1);sheet.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');SpreadsheetApp.flush();}
function runLotteonRemoteTaskContinueRemote_(){return runLotteonRemoteTaskStartRemote_();}
