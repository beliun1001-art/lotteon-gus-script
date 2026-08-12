/**
 * Issue #51 v0.1 read-only preflight.
 * Checks whether the newly rebuilt 2026 Apr-Jun VAT order set can be safely
 * passed into the existing v6.64/v6.69/v6.70 card matching chain.
 * Writes only ISSUE51_* diagnostic sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE51-v0.1-20260812',
  title: '현재 VAT 1,893주문 카드매칭 사전점검',
  enabled: true,
  outputSheet: 'ISSUE51_카드매칭사전점검',
  statusSheet: 'ISSUE51_실행상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var status = issue51Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue51Write_(status,[
    ['항목','값'],['버전','v0.1-ISSUE51-CARD-REMATCH-PREFLIGHT'],
    ['상태','RUNNING'],['단계','LOAD'],['메시지','현재 VAT ↔ 기존 카드검증/증빙 사전점검 시작'],['운영시트 변경','0']
  ]);
  try {
    var vat = ss.getSheetByName('부가세_신고자료');
    var verify = ss.getSheetByName('부가세_카드매칭검증');
    var history = ss.getSheetByName('카드사용내역_붙여넣기');
    var master = ss.getSheetByName('카드_마스터');
    var source = ss.getSheetByName('매출데이터_붙여넣기');
    if (!vat || vat.getLastRow() < 2) throw new Error('부가세_신고자료가 없습니다.');
    if (!verify || verify.getLastRow() < 2) throw new Error('부가세_카드매칭검증이 없습니다.');
    if (!history || history.getLastRow() < 2) throw new Error('카드사용내역_붙여넣기가 없습니다.');
    if (!master || master.getLastRow() < 2) throw new Error('카드_마스터가 없습니다.');

    var vv = vat.getDataRange().getValues(), vh = vv[0].map(issue51Text_);
    var xi = issue51HeaderIndexes_(vh, {
      year:['신고연도'], half:['반기'], account:['쿠팡계정ID'], order:['주문번호'], purchase:['매입금액'], payment:['롯데결제수단']
    });
    issue51Require_(xi.year>=0 && xi.half>=0 && xi.account>=0 && xi.order>=0 && xi.purchase>=0, 'VAT 필수 헤더 누락');

    var vatMap={}, vatRawByKey={}, vatPurchase=0, vatRows=0;
    for (var r=1;r<vv.length;r++) {
      var row=vv[r];
      if (issue51Text_(row[xi.year])!=='2026' || issue51Text_(row[xi.half])!=='상반기') continue;
      vatRows++;
      var account=issue51Text_(row[xi.account]).toLowerCase();
      var rawOrder=issue51Text_(row[xi.order]);
      var norm=issue51NormOrder_(rawOrder);
      if (!account || !norm) continue;
      var key=account+'|'+norm;
      if (!vatMap[key]) vatMap[key]={account:account,norm:norm,raw:rawOrder,purchase:0,rows:0};
      vatMap[key].purchase += issue51Num_(row[xi.purchase]); vatMap[key].rows++;
      vatPurchase += issue51Num_(row[xi.purchase]);
      if (!vatRawByKey[key]) vatRawByKey[key]={}; vatRawByKey[key][rawOrder]=true;
    }
    var vatKeys=Object.keys(vatMap), vatNormCollisions=0;
    vatKeys.forEach(function(k){if(Object.keys(vatRawByKey[k]||{}).length>1)vatNormCollisions++;});

    var qv = verify.getDataRange().getValues(), qh=qv[0].map(issue51Text_);
    var qi=issue51HeaderIndexes_(qh,{
      year:['신고연도'],half:['반기'],account:['쿠팡계정ID'],order:['주문번호'],payment:['롯데결제수단'],status:['카드매칭상태'],purchase:['주문매입금액']
    });
    issue51Require_(qi.year>=0 && qi.half>=0 && qi.account>=0 && qi.order>=0, '기존 카드검증 필수 헤더 누락');
    var verifyMap={}, verifyDup=0, verifyRows=0;
    for (var i=1;i<qv.length;i++) {
      var qr=qv[i];
      if(issue51Text_(qr[qi.year])!=='2026'||issue51Text_(qr[qi.half])!=='상반기')continue;
      verifyRows++;
      var qa=issue51Text_(qr[qi.account]).toLowerCase(), qo=issue51NormOrder_(qr[qi.order]);
      if(!qa||!qo)continue;
      var qk=qa+'|'+qo;
      if(verifyMap[qk]) verifyDup++;
      else verifyMap[qk]={payment:qi.payment>=0?issue51Text_(qr[qi.payment]):'',status:qi.status>=0?issue51Text_(qr[qi.status]):'',purchase:qi.purchase>=0?issue51Num_(qr[qi.purchase]):0};
    }
    var verifyKeys=Object.keys(verifyMap), overlap=0, currentOnly=0, verifyOnly=0, overlapPaymentNonBlank=0, overlapPaymentBlank=0;
    vatKeys.forEach(function(k){
      if(verifyMap[k]){overlap++; if(verifyMap[k].payment)overlapPaymentNonBlank++;else overlapPaymentBlank++;}
      else currentOnly++;
    });
    verifyKeys.forEach(function(k){if(!vatMap[k])verifyOnly++;});

    var sh = history.getDataRange().getValues(), hh=sh[0].map(issue51Text_);
    var hiDate=issue51FindHeader_(hh,['승인일','이용일','거래일','사용일','승인일자']);
    var histH1=0, firstDate='', lastDate='';
    for(var h=1;h<sh.length;h++){
      var d=hiDate>=0?issue51Date_(sh[h][hiDate]):'';
      if(d && d>='2026-01-01' && d<='2026-06-30'){histH1++;if(!firstDate||d<firstDate)firstDate=d;if(!lastDate||d>lastDate)lastDate=d;}
    }

    var sourcePayment='없음';
    if(source && source.getLastRow()>=1){
      var srcH=source.getRange(1,1,1,source.getLastColumn()).getValues()[0].map(issue51Text_);
      var spi=issue51FindHeader_(srcH,['롯데결제수단','결제수단','결제정보','결제방법','구매결제수단','결제수단/카드사','결제수단(카드사)']);
      if(spi>=0) sourcePayment=issue51Col_(spi+1)+' / '+srcH[spi];
    }

    var outRows=[
      ['구분','값','판정'],
      ['현재VAT 주문',vatKeys.length,vatKeys.length===1893?'OK':'CHECK'],
      ['현재VAT 상세행',vatRows,vatRows===3894?'OK':'CHECK'],
      ['현재VAT 매입합계',Math.round(vatPurchase),Math.round(vatPurchase)===106707957?'OK':'CHECK'],
      ['현재VAT 롯데결제수단열',xi.payment>=0?issue51Col_(xi.payment+1)+' / '+vh[xi.payment]:'없음',xi.payment>=0?'AVAILABLE':'MISSING'],
      ['원천 결제수단열',sourcePayment,sourcePayment==='없음'?'MISSING':'AVAILABLE'],
      ['기존검증 주문',verifyKeys.length,'INFO'],
      ['현재VAT↔기존검증 겹침',overlap,'INFO'],
      ['현재VAT 신규주문',currentOnly,'INFO'],
      ['기존검증에만 존재',verifyOnly,'INFO'],
      ['겹침 중 기존 결제수단 있음',overlapPaymentNonBlank,'INFO'],
      ['겹침 중 기존 결제수단 공란',overlapPaymentBlank,'INFO'],
      ['기존검증 정규화키 중복행',verifyDup,verifyDup===0?'OK':'CHECK'],
      ['현재VAT 정규화키 표현충돌',vatNormCollisions,vatNormCollisions===0?'OK':'CHECK'],
      ['카드사용내역 전체행',Math.max(0,history.getLastRow()-1),'INFO'],
      ['카드사용내역 H1행',histH1,'INFO'],
      ['카드사용내역 H1기간',(firstDate||'')+'~'+(lastDate||''),'INFO'],
      ['카드마스터행',Math.max(0,master.getLastRow()-1),'INFO']
    ];
    var out=issue51Ensure_(ss,LOTTEON_REMOTE_TASK.outputSheet); issue51WriteN_(out,outRows);

    var readyExact = vatKeys.length===1893 && vatRows===3894 && Math.round(vatPurchase)===106707957 && verifyDup===0 && vatNormCollisions===0 && histH1>0;
    var paymentCoverage = overlapPaymentNonBlank;
    var s=[
      ['항목','값'],['버전','v0.1-ISSUE51-CARD-REMATCH-PREFLIGHT'],['상태','PASS'],['단계','DONE'],
      ['메시지','현재 VAT ↔ 기존 카드검증/증빙 사전점검 완료'],['운영시트 변경','0'],
      ['현재VAT상세행',vatRows],['현재VAT주문',vatKeys.length],['현재VAT매입합계',Math.round(vatPurchase)],
      ['현재VAT_롯데결제수단열',xi.payment>=0?issue51Col_(xi.payment+1)+' / '+vh[xi.payment]:'없음'],
      ['원천_결제수단열',sourcePayment],['기존검증주문',verifyKeys.length],['현재VAT_기존검증겹침',overlap],
      ['현재VAT신규주문',currentOnly],['기존검증에만존재',verifyOnly],
      ['겹침_기존결제수단있음',overlapPaymentNonBlank],['겹침_기존결제수단공란',overlapPaymentBlank],
      ['기존검증정규화키중복행',verifyDup],['현재VAT정규화표현충돌',vatNormCollisions],
      ['카드원본H1행',histH1],['카드원본H1기간',(firstDate||'')+'~'+(lastDate||'')],['카드마스터행',Math.max(0,master.getLastRow()-1)],
      ['정확금액매칭사전조건',readyExact?'PASS':'CHECK'],
      ['2·3차귀속결제수단재사용가능주문',paymentCoverage],
      ['완료시각',new Date().toISOString()]
    ];
    issue51Write_(status,s);
    return {ok:true,vatOrders:vatKeys.length,overlap:overlap,currentOnly:currentOnly,exactReady:readyExact};
  } catch(e) {
    issue51Write_(status,[['항목','값'],['버전','v0.1-ISSUE51-CARD-REMATCH-PREFLIGHT'],['상태','ERROR'],['단계','FAILED'],['메시지','카드매칭 사전점검 실패'],['오류',String(e&&e.message?e.message:e)],['운영시트 변경','0']]);
    throw e;
  }
}
function issue51HeaderIndexes_(h,spec){var o={};Object.keys(spec).forEach(function(k){o[k]=issue51FindHeader_(h,spec[k]);});return o;}
function issue51FindHeader_(h,names){for(var n=0;n<names.length;n++){var w=issue51Compact_(names[n]);for(var i=0;i<h.length;i++)if(issue51Compact_(h[i])===w)return i;}return -1;}
function issue51Text_(v){return String(v==null?'':v).trim();}
function issue51Compact_(v){return issue51Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function issue51NormOrder_(v){return issue51Text_(v).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');}
function issue51Num_(v){if(typeof v==='number'&&isFinite(v))return v;var n=Number(issue51Text_(v).replace(/[원,%\s,]/g,''));return isFinite(n)?n:0;}
function issue51Date_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');var s=issue51Text_(v),m=s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);if(m)return m[1]+'-'+issue51Pad_(m[2])+'-'+issue51Pad_(m[3]);return '';}
function issue51Pad_(v){v=String(v);return v.length<2?'0'+v:v;}
function issue51Col_(n){var s='';while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;}
function issue51Require_(ok,msg){if(!ok)throw new Error(msg);}
function issue51Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}
function issue51Write_(sh,rows){sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setFrozenRows(1);}
function issue51WriteN_(sh,rows){sh.clearContents();sh.getRange(1,1,rows.length,rows[0].length).setValues(rows);sh.getRange(1,1,1,rows[0].length).setFontWeight('bold');sh.setFrozenRows(1);}
