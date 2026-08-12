/**
 * Issue #48 standalone diagnostic.
 * Audits all 2026 H1 orders between current VAT detail and card verification
 * using normalized order-number JOIN keys. Reads production sheets and writes
 * only ISSUE48_* diagnostic sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE48-v1.0-20260812',
  title: '상반기 VAT 현재 상세 ↔ 카드매칭검증 1,355건 stale audit',
  enabled: true,
  inputSheet: 'ISSUE47_3단계JOIN진단',
  outputSheet: 'ISSUE48_VAT검증STALE감사',
  statusSheet: 'ISSUE48_진단상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var state = issue48Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue48Write_(state,[
    ['항목','값'],['버전','v1.0-ISSUE48-H1-VAT-VERIFY-STALE-AUDIT'],['상태','RUNNING'],['단계','LOAD'],
    ['메시지','현재 VAT 상세 ↔ 카드매칭검증 전체 stale audit 시작'],['운영시트 변경','0'],['갱신시각',new Date().toISOString()]
  ]);

  try {
    var vat = ss.getSheetByName('부가세_신고자료');
    var verify = ss.getSheetByName('부가세_카드매칭검증');
    var issue47 = ss.getSheetByName(LOTTEON_REMOTE_TASK.inputSheet);
    if (!vat || vat.getLastRow() < 2) throw new Error('부가세_신고자료 시트가 없습니다.');
    if (!verify || verify.getLastRow() < 2) throw new Error('부가세_카드매칭검증 시트가 없습니다.');
    if (!issue47 || issue47.getLastRow() < 2) throw new Error('ISSUE47_3단계JOIN진단 시트가 없습니다.');

    var targetSet = issue48TargetSet_(issue47);
    if (Object.keys(targetSet).length !== 5) throw new Error('Issue47 대상 건수 불일치: 기대 5건, 실제 ' + Object.keys(targetSet).length + '건');

    var vv = vat.getDataRange().getValues(), vh = vv[0] || [], vm = issue48Map_(vh);
    var vi = {
      year:issue48Find_(vm,['신고연도']), half:issue48Find_(vm,['반기']), business:issue48Find_(vm,['사업자등록번호']),
      account:issue48Find_(vm,['쿠팡계정ID','마켓아이디','계정ID']), order:issue48Find_(vm,['주문번호','마켓주문번호','주문ID','주문ID(마켓)']),
      date:issue48Find_(vm,['날짜','주문일','주문일자','마켓주문일자']), payment:issue48Find_(vm,['롯데결제수단','구매결제수단','결제수단']),
      purchase:issue48Find_(vm,['매입금액','주문매입금액','구매가격','매입가격','상품매입금액'])
    };
    ['year','half','business','account','order','purchase'].forEach(function(k){if(vi[k]<0) throw new Error('VAT 필수 헤더 누락: '+k);});

    var cv = verify.getDataRange().getValues(), ch = cv[0] || [], cm = issue48Map_(ch);
    var ci = {
      year:issue48Find_(cm,['신고연도']), half:issue48Find_(cm,['반기']), business:issue48Find_(cm,['사업자등록번호']),
      account:issue48Find_(cm,['쿠팡계정ID','마켓아이디','계정ID']), order:issue48Find_(cm,['주문번호','마켓주문번호','주문ID','주문ID(마켓)']),
      date:issue48Find_(cm,['주문일','날짜','주문일자']), payment:issue48Find_(cm,['롯데결제수단','구매결제수단','결제수단']),
      purchase:issue48Find_(cm,['주문매입금액','매입금액']), status:issue48Find_(cm,['카드매칭상태'])
    };
    ['year','half','business','account','order','purchase','status'].forEach(function(k){if(ci[k]<0) throw new Error('검증 필수 헤더 누락: '+k);});

    var vatOrders = issue48GroupVat_(vv, vi);
    var verifyOrders = issue48ReadVerify_(cv, ci);
    var vatMap = issue48Bucket_(vatOrders), verifyMap = issue48Bucket_(verifyOrders);
    var allKeys = {};
    Object.keys(vatMap).forEach(function(k){allKeys[k]=true;});
    Object.keys(verifyMap).forEach(function(k){allKeys[k]=true;});

    var headers = [
      'JOIN키','사업자등록번호','쿠팡계정ID','정규화주문번호','VAT주문번호','검증주문번호','VAT주문행수','검증주문행수',
      'VAT상세행수','VAT매입금액','검증주문매입금액','차액(VAT-검증)','카드매칭상태','분류','Issue47대상','VAT문장부호','검증문장부호','메모'
    ];
    var out = [];
    var stats = {
      vatOrders:vatOrders.length, verifyOrders:verifyOrders.length, oneToOne:0, vatOnly:0, verifyOnly:0, duplicateKeys:0,
      amountMatch:0, amountMismatch:0, vatPosVerifyZero:0, vatZeroVerifyPos:0, otherDiff:0,
      vatPurchase:0, verifyPurchase:0, targetCovered:0, targetVatPosVerifyZero:0
    };
    var statusCounts = {}, mismatchStatusCounts = {}, targetSeen = {};

    vatOrders.forEach(function(x){ stats.vatPurchase += x.purchase; });
    verifyOrders.forEach(function(x){
      stats.verifyPurchase += x.purchase;
      var s = x.status || '공란'; statusCounts[s]=(statusCounts[s]||0)+1;
    });

    Object.keys(allKeys).sort().forEach(function(key){
      var va = vatMap[key] || [], ca = verifyMap[key] || [];
      var classification = '', vp = 0, cp = 0, status = '', vatRaw = '', verifyRaw = '', vatDetailRows = 0;
      va.forEach(function(x){vp += x.purchase; vatDetailRows += x.detailRows;});
      ca.forEach(function(x){cp += x.purchase;});
      if (va.length) vatRaw = issue48Distinct_(va.map(function(x){return x.orderRaw;})).join(' | ');
      if (ca.length) verifyRaw = issue48Distinct_(ca.map(function(x){return x.orderRaw;})).join(' | ');
      if (ca.length) status = issue48Distinct_(ca.map(function(x){return x.status;})).join(' | ');

      if (!va.length) { classification='검증_ONLY'; stats.verifyOnly++; }
      else if (!ca.length) { classification='VAT_ONLY'; stats.vatOnly++; }
      else if (va.length!==1 || ca.length!==1) { classification='정규화키중복'; stats.duplicateKeys++; }
      else {
        stats.oneToOne++;
        if (Math.round(vp) === Math.round(cp)) { classification='매입금액일치'; stats.amountMatch++; }
        else {
          stats.amountMismatch++;
          if (vp>0 && cp===0) { classification='VAT양수_검증0'; stats.vatPosVerifyZero++; }
          else if (vp===0 && cp>0) { classification='VAT0_검증양수'; stats.vatZeroVerifyPos++; }
          else { classification='기타금액차이'; stats.otherDiff++; }
          var ms = status || '공란'; mismatchStatusCounts[ms]=(mismatchStatusCounts[ms]||0)+1;
        }
      }

      var sample = va.length ? va[0] : ca[0];
      var tkey = issue48TargetKey_(sample.account, sample.orderNorm);
      var isTarget = !!targetSet[tkey];
      if (isTarget && !targetSeen[tkey]) {
        targetSeen[tkey]=true; stats.targetCovered++;
        if (classification==='VAT양수_검증0') stats.targetVatPosVerifyZero++;
      }
      var memo=[];
      if(va.length!==1) memo.push('VAT주문그룹='+va.length);
      if(ca.length!==1) memo.push('검증행='+ca.length);
      if(vatRaw && verifyRaw && vatRaw!==verifyRaw && issue48OrderNorm_(vatRaw)===issue48OrderNorm_(verifyRaw)) memo.push('표현만다름');
      out.push([
        key,sample.business,sample.account,sample.orderNorm,vatRaw,verifyRaw,va.length,ca.length,vatDetailRows,vp,cp,vp-cp,status,classification,isTarget?'Y':'N',
        issue48Punct_(vatRaw),issue48Punct_(verifyRaw),memo.join(' / ')
      ]);
    });

    var output = issue48Ensure_(ss, LOTTEON_REMOTE_TASK.outputSheet);
    output.clearContents();
    output.getRange(1,1,1,headers.length).setValues([headers]);
    if(out.length) output.getRange(2,1,out.length,headers.length).setValues(out);
    output.setFrozenRows(1);
    output.getRange(1,1,1,headers.length).setBackground('#d9eaf7').setFontWeight('bold');
    if(out.length) output.getRange(2,10,out.length,3).setNumberFormat('#,##0');

    var statusRows = [
      ['항목','값'],['버전','v1.0-ISSUE48-H1-VAT-VERIFY-STALE-AUDIT'],['상태','PASS'],['단계','DONE'],
      ['메시지','현재 VAT 상세 ↔ 카드매칭검증 전체 stale audit 완료'],['운영시트 변경','0'],
      ['VAT상반기주문',stats.vatOrders],['검증상반기주문',stats.verifyOrders],['정규화1대1매칭',stats.oneToOne],
      ['VAT_ONLY',stats.vatOnly],['검증_ONLY',stats.verifyOnly],['정규화키중복',stats.duplicateKeys],
      ['매입금액일치',stats.amountMatch],['매입금액불일치',stats.amountMismatch],['VAT양수_검증0',stats.vatPosVerifyZero],
      ['VAT0_검증양수',stats.vatZeroVerifyPos],['기타금액차이',stats.otherDiff],
      ['VAT총매입금액',Math.round(stats.vatPurchase)],['검증총매입금액',Math.round(stats.verifyPurchase)],['총매입금액차액',Math.round(stats.vatPurchase-stats.verifyPurchase)],
      ['Issue47대상5_감사포함',stats.targetCovered],['Issue47대상5_VAT양수검증0',stats.targetVatPosVerifyZero]
    ];
    Object.keys(statusCounts).sort().forEach(function(k){statusRows.push(['현재상태_'+k,statusCounts[k]]);});
    Object.keys(mismatchStatusCounts).sort().forEach(function(k){statusRows.push(['금액불일치상태_'+k,mismatchStatusCounts[k]]);});
    statusRows.push(['완료시각',new Date().toISOString()]);
    issue48Write_(state,statusRows);
    try{MailApp.sendEmail('beliun1001@gmail.com','[LOTTEON 자동작업 결과][PASS] ISSUE48-v1.0',statusRows.map(function(x){return x[0]+': '+x[1];}).join('\n'));}catch(mailError){}
    return {ok:true,stats:stats};
  } catch(e) {
    issue48Write_(state,[
      ['항목','값'],['버전','v1.0-ISSUE48-H1-VAT-VERIFY-STALE-AUDIT'],['상태','ERROR'],['단계','FAILED'],
      ['메시지','현재 VAT 상세 ↔ 카드매칭검증 stale audit 실패'],['오류',String(e&&e.message?e.message:e)],['운영시트 변경','0'],['갱신시각',new Date().toISOString()]
    ]);
    throw e;
  }
}

function issue48TargetSet_(sheet){
  var v=sheet.getDataRange().getValues(), m=issue48Map_(v[0]||[]), ni=issue48Find_(m,['정규화주문번호']), ai=issue48Find_(m,['쿠팡계정ID']);
  if(ni<0||ai<0) throw new Error('ISSUE47 대상 헤더 누락');
  var out={}; for(var r=1;r<v.length;r++){var n=issue48Text_(v[r][ni]),a=issue48Text_(v[r][ai]);if(n)out[issue48TargetKey_(a,n)]=true;} return out;
}
function issue48GroupVat_(values,ix){
  var exact={};
  for(var r=1;r<values.length;r++){
    var row=values[r], year=issue48Text_(row[ix.year]), half=issue48Text_(row[ix.half]);
    if(year!=='2026'||half!=='상반기') continue;
    var business=issue48Text_(row[ix.business]), account=issue48Text_(row[ix.account]), orderRaw=issue48Text_(row[ix.order]);
    var date=ix.date>=0?issue48Date_(row[ix.date]):'', payment=ix.payment>=0?issue48Text_(row[ix.payment]):'';
    var exactKey=orderRaw?[year,half,business,account,orderRaw].join('|'):[year,half,business,account,'BLANK',date,r].join('|');
    if(!exact[exactKey]) exact[exactKey]={business:business,account:account,orderRaw:orderRaw,orderNorm:issue48OrderNorm_(orderRaw),date:date,payments:{},purchase:0,detailRows:0};
    exact[exactKey].purchase+=issue48Number_(row[ix.purchase]); exact[exactKey].detailRows++;
    if(payment) exact[exactKey].payments[payment]=true;
  }
  return Object.keys(exact).map(function(k){var x=exact[k];x.payment=Object.keys(x.payments).sort().join(', ');x.joinKey=issue48JoinKey_(x);return x;});
}
function issue48ReadVerify_(values,ix){
  var out=[];
  for(var r=1;r<values.length;r++){
    var row=values[r], year=issue48Text_(row[ix.year]), half=issue48Text_(row[ix.half]);
    if(year!=='2026'||half!=='상반기') continue;
    var x={business:issue48Text_(row[ix.business]),account:issue48Text_(row[ix.account]),orderRaw:issue48Text_(row[ix.order]),date:ix.date>=0?issue48Date_(row[ix.date]):'',payment:ix.payment>=0?issue48Text_(row[ix.payment]):'',purchase:issue48Number_(row[ix.purchase]),status:issue48Text_(row[ix.status]),detailRows:1};
    x.orderNorm=issue48OrderNorm_(x.orderRaw); x.joinKey=issue48JoinKey_(x); out.push(x);
  }
  return out;
}
function issue48JoinKey_(x){
  var base=['2026','상반기',issue48Text_(x.business),issue48Text_(x.account)];
  if(x.orderNorm) return base.concat(['ORDER',x.orderNorm]).join('|');
  return base.concat(['BLANK',issue48Text_(x.date),issue48Compact_(x.payment)]).join('|');
}
function issue48Bucket_(items){var m={};(items||[]).forEach(function(x){if(!m[x.joinKey])m[x.joinKey]=[];m[x.joinKey].push(x);});return m;}
function issue48TargetKey_(account,norm){return issue48Text_(account).toLowerCase()+'|'+issue48Text_(norm).toLowerCase();}
function issue48Distinct_(a){var m={},o=[];(a||[]).forEach(function(x){x=issue48Text_(x);if(!m[x]){m[x]=true;o.push(x);}});return o;}
function issue48Punct_(s){s=issue48Text_(s);var a=[];for(var i=0;i<s.length;i++){var c=s.charAt(i);if(!/[0-9a-zA-Z가-힣]/.test(c))a.push(c+'(U+'+('0000'+c.charCodeAt(0).toString(16).toUpperCase()).slice(-4)+')');}return a.join(' ');}
function issue48OrderNorm_(v){return issue48Text_(v).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');}
function issue48Compact_(v){return issue48Text_(v).toLowerCase().replace(/\s+/g,'').replace(/[^0-9a-z가-힣]/g,'');}
function issue48Map_(h){var m={};(h||[]).forEach(function(x,i){m[issue48Norm_(x)]=i;});return m;}
function issue48Find_(m,a){for(var i=0;i<a.length;i++){var k=issue48Norm_(a[i]);if(m[k]!=null)return m[k];}return -1;}
function issue48Norm_(v){return issue48Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function issue48Text_(v){return v==null?'':String(v).trim();}
function issue48Number_(v){if(typeof v==='number'&&isFinite(v))return v;var n=Number(String(v==null?'':v).replace(/,/g,'').replace(/[^0-9.\-]/g,''));return isFinite(n)?n:0;}
function issue48Date_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');return issue48Text_(v);}
function issue48Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}
function issue48Write_(s,r){s.clearContents();s.getRange(1,1,r.length,2).setValues(r);s.setFrozenRows(1);s.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');SpreadsheetApp.flush();}
function runLotteonRemoteTaskContinueRemote_(){return runLotteonRemoteTaskStartRemote_();}
