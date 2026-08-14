var LOTTEON_REMOTE_TASK={id:'ISSUE72-v1.0-20260814',title:'NON_CARD 498 현금영수증 분류 + NO_MATCH 49 증빙원인 진단',enabled:true,statusSheet:'ISSUE72_실행상태'};
var I72_VERSION='v1.0-ISSUE72-NONCARD-CASHRECEIPT-NOMATCH-DIAGNOSTIC';
var I72_DETAIL='ISSUE72_증빙분류';
var I72_STATUS='ISSUE72_실행상태';
var I72_CORE=['매출데이터_붙여넣기','부가세_신고자료','부가세_카드매칭검증','부가세_기간별','카드사용내역_붙여넣기','카드_마스터'];
var I72_OLD_TEMP_RE=/^ISSUE(?:5[3-9]|6[0-8])(?:_|$)/i;

function runLotteonRemoteTaskStartRemote_(){return i72Run_();}
function runLotteonRemoteTaskContinueRemote_(){return i72Run_();}

function i72Run_(){
  var ss=SpreadsheetApp.getActive();
  var before={};
  try{
    i72Status_(ss,[['version',I72_VERSION],['상태','RUNNING'],['단계','GUARD'],['메시지','운영 원본 read-only 가드 확인 중'],['갱신시각',new Date().toISOString()]]);

    var missing=[];
    I72_CORE.forEach(function(name){var sh=ss.getSheetByName(name);if(!sh)missing.push(name);else before[name]=i72Sig_(sh);});
    if(missing.length) throw new Error('핵심 시트 누락: '+missing.join(', '));

    var oldTemps=ss.getSheets().map(function(sh){return sh.getName();}).filter(function(n){return I72_OLD_TEMP_RE.test(n);});
    if(oldTemps.length) throw new Error('Issue71 정리 미완료: ISSUE53~68 임시시트 잔여 '+oldTemps.join(', '));

    var cardSheet=ss.getSheetByName('부가세_카드매칭검증');
    var card=i72ReadCardVerify_(cardSheet);
    i72GuardCard_(card);

    var hist=i72ReadHistory_(ss.getSheetByName('카드사용내역_붙여넣기'));
    var canonical=i72CanonicalizeHistory_(hist);

    var out=[];
    var nc={confirmedCount:0,confirmedPurchase:0,kakaoReviewCount:0,kakaoReviewPurchase:0,otherCount:0,otherPurchase:0};
    var nm={zero:0,cash:0,card:0,canceled:0,date:0,amount:0,none:0,other:0};
    var nmPurchase={zero:0,cash:0,card:0,canceled:0,date:0,amount:0,none:0,other:0};
    var business={};
    var payment={};

    card.rows.forEach(function(o){
      if(o.status!=='NON_CARD' && o.status!=='NO_MATCH') return;
      if(!business[o.business]) business[o.business]={nonCard:0,nonCardPurchase:0,cashEligible:0,cashEligiblePurchase:0,noMatch:0,noMatchPurchase:0};
      var b=business[o.business];
      var diag={exactCard:0,exactNonCard:0,fullCanceled:0,outsideExact:0,periodAmountDiff:0,nearestDiff:'',candidateSummary:''};
      var reportType='', internal='', noMatchCause='';

      if(o.status==='NON_CARD'){
        b.nonCard++;b.nonCardPurchase+=o.purchase;
        var cls=i72ClassifyNonCard_(o);
        reportType=cls.reportType; internal=cls.internal;
        if(internal==='현금영수증_확정'){nc.confirmedCount++;nc.confirmedPurchase+=o.purchase;b.cashEligible++;b.cashEligiblePurchase+=o.purchase;}
        else if(internal==='카카오페이머니_현금영수증확인필요'){nc.kakaoReviewCount++;nc.kakaoReviewPurchase+=o.purchase;}
        else {nc.otherCount++;nc.otherPurchase+=o.purchase;}
      }else{
        b.noMatch++;b.noMatchPurchase+=o.purchase;
        var pKind=i72PaymentKind_(o.lottePayment);
        payment[pKind]=(payment[pKind]||0)+1;
        var d=i72DiagnoseNoMatch_(o,canonical,hist);
        diag=d;
        noMatchCause=d.cause;
        reportType='미확정'; internal='NO_MATCH';
        var nk=i72NoMatchBucket_(d.cause);
        nm[nk]++;nmPurchase[nk]+=o.purchase;
      }

      out.push([
        o.status==='NON_CARD'?'NON_CARD':'NO_MATCH',reportType,internal,o.business,o.account,o.orderDate,o.orderNo,o.purchase,o.lottePayment,o.status,
        o.company,o.alias,o.cardName,o.evidenceType,o.reason,o.sourceFile,
        diag.exactCard,diag.exactNonCard,diag.fullCanceled,diag.outsideExact,diag.periodAmountDiff,diag.nearestDiff,noMatchCause,diag.candidateSummary
      ]);
    });

    if(out.length!==547) throw new Error('진단 대상 행수 불일치: '+out.length+' (기대 547)');
    if(nc.confirmedCount+nc.kakaoReviewCount+nc.otherCount!==498) throw new Error('NON_CARD 분류 합계 불일치');
    if(nm.zero+nm.cash+nm.card+nm.canceled+nm.date+nm.amount+nm.none+nm.other!==49) throw new Error('NO_MATCH 분류 합계 불일치');

    i72WriteDetail_(ss,out);

    var changed=[];
    I72_CORE.forEach(function(name){var sh=ss.getSheetByName(name);if(!sh||i72Sig_(sh)!==before[name])changed.push(name);});
    if(changed.length) throw new Error('보호 핵심 시트 변경 감지: '+changed.join(', '));

    var status=[
      ['version',I72_VERSION],['상태','PASS'],['단계','DONE'],['메시지','NON_CARD 498 현금영수증 확정/확인필요 분류 및 NO_MATCH 49 증빙원인 read-only 진단 완료'],
      ['cleanup잔여임시시트수',0],['카드검증_주문수',card.rows.length],['MATCHED',card.stats.MATCHED||0],['NON_CARD',card.stats.NON_CARD||0],['AMBIGUOUS',card.stats.AMBIGUOUS||0],['NO_MATCH',card.stats.NO_MATCH||0],['매입합계',card.purchase],
      ['NONCARD_현금영수증확정건수',nc.confirmedCount],['NONCARD_현금영수증확정_매입합계',nc.confirmedPurchase],
      ['NONCARD_카카오페이머니_현금영수증확인필요건수',nc.kakaoReviewCount],['NONCARD_카카오페이머니_확인필요_매입합계',nc.kakaoReviewPurchase],
      ['NONCARD_기타비카드건수',nc.otherCount],['NONCARD_기타비카드_매입합계',nc.otherPurchase],
      ['NOMATCH_매입금액0',nm.zero],['NOMATCH_현금성증빙후보있음',nm.cash],['NOMATCH_카드증빙후보있음',nm.card],['NOMATCH_완전취소증빙만',nm.canceled],['NOMATCH_날짜차이정확금액',nm.date],['NOMATCH_기간내금액불일치',nm.amount],['NOMATCH_증빙자체없음',nm.none],['NOMATCH_기타',nm.other],
      ['NOMATCH_매입금액0_매입합계',nmPurchase.zero],['NOMATCH_현금성후보_매입합계',nmPurchase.cash],['NOMATCH_카드후보_매입합계',nmPurchase.card],['NOMATCH_완전취소_매입합계',nmPurchase.canceled],['NOMATCH_날짜차이_매입합계',nmPurchase.date],['NOMATCH_금액불일치_매입합계',nmPurchase.amount],['NOMATCH_증빙없음_매입합계',nmPurchase.none],['NOMATCH_기타_매입합계',nmPurchase.other]
    ];
    Object.keys(payment).sort().forEach(function(k){status.push(['NOMATCH_결제수단_'+k,payment[k]]);});
    Object.keys(business).sort().forEach(function(k){var x=business[k];status.push(['사업자_'+k+'_NONCARD',x.nonCard]);status.push(['사업자_'+k+'_현금영수증확정',x.cashEligible]);status.push(['사업자_'+k+'_NO_MATCH',x.noMatch]);status.push(['사업자_'+k+'_NO_MATCH매입',x.noMatchPurchase]);});
    status.push(['보호시트변경수',0],['오류',''],['완료시각',new Date().toISOString()]);
    i72Status_(ss,status);
    try{ss.toast('Issue72 PASS: 현금영수증 확정 '+nc.confirmedCount+'건 / 카카오머니 확인필요 '+nc.kakaoReviewCount+'건 / NO_MATCH 49건 진단 완료','LOTTEON',10);}catch(_e){}
    return {ok:true,done:true,version:I72_VERSION,cashReceiptConfirmed:nc.confirmedCount,kakaoMoneyReview:nc.kakaoReviewCount,noMatch:49,coreChanged:0};
  }catch(e){
    var msg=String(e&&e.message?e.message:e);
    try{i72Status_(ss,[['version',I72_VERSION],['상태','ERROR'],['단계','FAILED'],['메시지','Issue72 진단 실패'],['오류',msg],['완료시각',new Date().toISOString()]]);}catch(_e2){}
    try{ss.toast('Issue72 실패: '+msg,'LOTTEON',10);}catch(_e3){}
    throw e;
  }
}

function i72ReadCardVerify_(sh){
  var values=sh.getDataRange().getValues();
  var hr=i72FindHeaderRow_(values,['주문번호','카드매칭상태']);
  if(hr<0) throw new Error('부가세_카드매칭검증 header 탐지 실패');
  var h=values[hr];
  var ix={
    date:i72Ix_(h,['주문일']), business:i72Ix_(h,['사업자등록번호']), account:i72Ix_(h,['쿠팡계정ID','사업자코드']), order:i72Ix_(h,['주문번호']), payment:i72Ix_(h,['롯데결제수단']), purchase:i72Ix_(h,['주문매입금액','매입금액']),
    company:i72Ix_(h,['구매카드사']), alias:i72Ix_(h,['구매카드별칭']), cardName:i72Ix_(h,['구매카드명']), status:i72Ix_(h,['카드매칭상태']), reason:i72Ix_(h,['카드매칭근거']), evidence:i72Ix_(h,['증빙유형']), source:i72Ix_(h,['원본파일'])
  };
  ['date','business','account','order','payment','purchase','company','alias','cardName','status','reason','evidence','source'].forEach(function(k){if(ix[k]<0 && ['payment','company','alias','cardName','reason','evidence','source'].indexOf(k)<0)throw new Error('카드검증 필수 header 누락: '+k);});
  var rows=[], stats={}, purchase=0, keys={};
  for(var r=hr+1;r<values.length;r++){
    var row=values[r], order=i72Text_(row[ix.order]); if(!order)continue;
    var account=i72Text_(row[ix.account]); var key=i72Compact_(account)+'|'+i72Compact_(order);
    if(keys[key]) throw new Error('카드검증 주문키 중복: '+key); keys[key]=true;
    var status=i72Text_(row[ix.status]).toUpperCase(); var p=i72Num_(row[ix.purchase]);
    stats[status]=(stats[status]||0)+1;purchase+=p;
    rows.push({rowNo:r+1,orderDate:i72Date_(row[ix.date]),business:i72Text_(row[ix.business]),account:account,orderNo:order,lottePayment:ix.payment>=0?i72Text_(row[ix.payment]):'',purchase:p,status:status,company:ix.company>=0?i72Text_(row[ix.company]):'',alias:ix.alias>=0?i72Text_(row[ix.alias]):'',cardName:ix.cardName>=0?i72Text_(row[ix.cardName]):'',reason:ix.reason>=0?i72Text_(row[ix.reason]):'',evidenceType:ix.evidence>=0?i72Text_(row[ix.evidence]):'',sourceFile:ix.source>=0?i72Text_(row[ix.source]):''});
  }
  return {rows:rows,stats:stats,purchase:Math.round(purchase)};
}

function i72GuardCard_(x){
  if(x.rows.length!==1355) throw new Error('카드검증 주문수 불일치: '+x.rows.length);
  if((x.stats.MATCHED||0)!==808||(x.stats.NON_CARD||0)!==498||(x.stats.AMBIGUOUS||0)!==0||(x.stats.NO_MATCH||0)!==49) throw new Error('카드검증 상태 집계 불일치: '+JSON.stringify(x.stats));
  if(x.purchase!==105762969) throw new Error('카드검증 매입합계 불일치: '+x.purchase);
}

function i72ReadHistory_(sh){
  var v=sh.getDataRange().getValues(); if(v.length<2)return [];
  var hr=i72FindHeaderRow_(v,['카드사','승인금액']); if(hr<0)hr=0; var h=v[hr];
  var ix={company:i72Ix_(h,['카드사']),name:i72Ix_(h,['카드명']),number:i72Ix_(h,['카드번호']),end4:i72Ix_(h,['카드번호끝4']),date:i72Ix_(h,['승인일','이용일','거래일']),time:i72Ix_(h,['승인시각','이용시각','거래시각']),merchant:i72Ix_(h,['가맹점명','이용가맹점']),amount:i72Ix_(h,['승인금액','이용금액','거래금액']),approval:i72Ix_(h,['승인번호']),status:i72Ix_(h,['승인상태','승인/취소구분','상태']),cancelDate:i72Ix_(h,['취소일']),cancelAmount:i72Ix_(h,['취소금액']),orderNo:i72Ix_(h,['가맹점주문번호','주문번호']),evidence:i72Ix_(h,['증빙유형']),lotte:i72Ix_(h,['롯데계열여부']),source:i72Ix_(h,['원본파일']),memo:i72Ix_(h,['메모'])};
  if(ix.date<0||ix.amount<0) throw new Error('카드사용내역 필수 header 누락');
  var out=[];
  for(var r=hr+1;r<v.length;r++){
    var row=v[r]; var date=i72Date_(row[ix.date]); var amount=i72Num_(row[ix.amount]); var orderNo=ix.orderNo>=0?i72Text_(row[ix.orderNo]):'';
    if(!date&&!amount&&!orderNo)continue;
    var o={rowNo:r+1,company:ix.company>=0?i72Text_(row[ix.company]):'',cardName:ix.name>=0?i72Text_(row[ix.name]):'',date:date,time:ix.time>=0?i72Text_(row[ix.time]):'',merchant:ix.merchant>=0?i72Text_(row[ix.merchant]):'',amount:amount,approvalNo:ix.approval>=0?i72Text_(row[ix.approval]):'',status:ix.status>=0?i72Text_(row[ix.status]):'',cancelDate:ix.cancelDate>=0?i72Date_(row[ix.cancelDate]):'',cancelAmount:ix.cancelAmount>=0?i72Num_(row[ix.cancelAmount]):0,merchantOrderNo:orderNo,evidenceType:ix.evidence>=0?i72Text_(row[ix.evidence]):'',lotteFlag:ix.lotte>=0?i72Text_(row[ix.lotte]):'',sourceFile:ix.source>=0?i72Text_(row[ix.source]):'',memo:ix.memo>=0?i72Text_(row[ix.memo]):''};
    o.nonCard=i72IsNonCardEvidence_(o); o.lotteEvidence=i72IsLotteEvidence_(o); o.cancelLike=i72IsCancel_(o);
    out.push(o);
  }
  return out;
}

function i72CanonicalizeHistory_(rows){
  var groups={}, singles=[];
  rows.forEach(function(h){
    var issuer=i72Compact_(h.company), approval=i72Compact_(h.approvalNo);
    if(!h.nonCard&&issuer&&approval){var k='CARD|'+issuer+'|'+approval;if(!groups[k])groups[k]=[];groups[k].push(h);}else{var s=i72CanonicalSingle_(h);if(s)singles.push(s);}
  });
  Object.keys(groups).forEach(function(k){var a=groups[k], positives=a.filter(function(x){return x.amount>0&&!x.cancelLike;});if(!positives.length)return;var rep=positives[0], original=0,cancel=0;a.forEach(function(x){if(x.amount>original)original=x.amount;var c=Math.abs(x.cancelAmount||0);if(x.amount<0)c=Math.max(c,Math.abs(x.amount));if(x.cancelLike&&x.amount>0)c=Math.max(c,x.amount);if(c>cancel)cancel=c;});cancel=Math.min(cancel,original);var eff=Math.max(original-cancel,0);singles.push(i72CanonObj_(rep,eff,original,eff===0));});
  return singles;
}
function i72CanonicalSingle_(h){if(h.cancelLike&&h.amount<0)return null;var original=Math.abs(h.amount||0),cancel=Math.abs(h.cancelAmount||0);if(h.cancelLike&&original>0)cancel=Math.max(cancel,original);var eff=Math.max(original-cancel,0);return i72CanonObj_(h,eff,original,original>0&&eff===0);}
function i72CanonObj_(h,eff,original,full){return {rowNo:h.rowNo,company:h.company,cardName:h.cardName,date:h.date,merchant:h.merchant,amount:eff,originalAmount:original,approvalNo:h.approvalNo,merchantOrderNo:h.merchantOrderNo,evidenceType:h.evidenceType,sourceFile:h.sourceFile,memo:h.memo,nonCard:h.nonCard,lotteEvidence:h.lotteEvidence,fullyCanceled:full};}

function i72ClassifyNonCard_(o){
  var s=i72Compact_([o.lottePayment,o.company,o.alias,o.cardName,o.evidenceType,o.reason,o.sourceFile].join(' '));
  if(s.indexOf('현금영수증')>=0) return {reportType:'현금영수증',internal:'현금영수증_확정'};
  if(i72Compact_(o.cardName).indexOf('카카오페이페이머니')>=0||i72Compact_(o.alias).indexOf('신한은행계좌결제')>=0||(s.indexOf('카카오')>=0&&(s.indexOf('페이머니')>=0||s.indexOf('머니')>=0||s.indexOf('계좌')>=0||s.indexOf('현금')>=0))) return {reportType:'현금영수증_확인필요',internal:'카카오페이머니_현금영수증확인필요'};
  return {reportType:'비카드_추가확인',internal:'기타_비카드'};
}

function i72DiagnoseNoMatch_(o,canon,raw){
  var d={exactCard:0,exactNonCard:0,fullCanceled:0,outsideExact:0,periodAmountDiff:0,nearestDiff:'',candidateSummary:'',cause:''};
  if(!o.purchase){d.cause='매입금액0';return d;}
  var exactCard=[],exactNon=[],canceled=[],outside=[],period=[];
  canon.forEach(function(h){
    if(!h.lotteEvidence||!h.date)return;
    var days=i72Days_(o.orderDate,h.date); var exact=Number(h.amount||0)===Number(o.purchase||0);
    if(days>=0&&days<=7){
      if(h.fullyCanceled&&Number(h.originalAmount||0)===Number(o.purchase||0))canceled.push(h);
      else if(!h.fullyCanceled&&exact){if(h.nonCard)exactNon.push(h);else exactCard.push(h);}
      else if(!h.fullyCanceled)period.push(h);
    } else if(days>=-14&&days<=14&&!h.fullyCanceled&&exact) outside.push(h);
  });
  d.exactCard=exactCard.length;d.exactNonCard=exactNon.length;d.fullCanceled=canceled.length;d.outsideExact=outside.length;d.periodAmountDiff=period.length;
  if(period.length){var nearest=period.slice().sort(function(a,b){return Math.abs(a.amount-o.purchase)-Math.abs(b.amount-o.purchase);})[0];d.nearestDiff=Math.abs(Number(nearest.amount||0)-Number(o.purchase||0));}
  if(exactNon.length&&exactCard.length)d.cause='기타_추가확인';
  else if(exactNon.length)d.cause='현금성증빙후보있음';
  else if(exactCard.length)d.cause='카드증빙후보있음';
  else if(canceled.length)d.cause='완전취소증빙만있음';
  else if(outside.length)d.cause='날짜차이_정확금액증빙있음';
  else if(period.length)d.cause='기간내_금액불일치증빙있음';
  else {
    var direct=raw.filter(function(h){return h.merchantOrderNo&&i72Compact_(h.merchantOrderNo)===i72Compact_(o.orderNo);});
    if(direct.length){d.cause='기타_추가확인';d.candidateSummary='가맹점주문번호 동일 원본증빙 '+direct.length+'건';}
    else d.cause='증빙자체없음';
  }
  if(!d.candidateSummary){var c=exactNon.concat(exactCard,canceled,outside,period).slice(0,3);d.candidateSummary=c.map(function(h){return [h.date,h.nonCard?'비카드':h.company,h.amount||h.originalAmount,h.merchant,h.evidenceType].join('|');}).join(' / ');}
  return d;
}
function i72NoMatchBucket_(c){if(c==='매입금액0')return 'zero';if(c==='현금성증빙후보있음')return 'cash';if(c==='카드증빙후보있음')return 'card';if(c==='완전취소증빙만있음')return 'canceled';if(c==='날짜차이_정확금액증빙있음')return 'date';if(c==='기간내_금액불일치증빙있음')return 'amount';if(c==='증빙자체없음')return 'none';return 'other';}
function i72PaymentKind_(v){var s=i72Compact_(v);if(!s)return '공란';if(s.indexOf('토스')>=0)return '토스페이';if(s.indexOf('카카오')>=0)return '카카오페이';if(s.indexOf('lpay')>=0||s.indexOf('엘페이')>=0)return 'L.PAY';if(/롯데|국민|kb|우리|신한|농협|nh|삼성|하나|현대/.test(s))return '카드사표기';return '기타';}

function i72WriteDetail_(ss,rows){
  var sh=ss.getSheetByName(I72_DETAIL)||ss.insertSheet(I72_DETAIL); sh.clear();
  var headers=['구분','신고증빙유형','내부분류','사업자등록번호','쿠팡계정ID','주문일','주문번호','주문매입금액','롯데결제수단','기존매칭상태','구매카드사','구매카드별칭','구매카드명','증빙유형','카드매칭근거','원본파일','정확0~+7_카드후보수','정확0~+7_비카드후보수','완전취소후보수','바깥기간정확금액후보수','동일기간금액다름후보수','최근접금액차이','NO_MATCH원인','후보요약'];
  sh.getRange(1,1,1,headers.length).setValues([headers]);if(rows.length)sh.getRange(2,1,rows.length,headers.length).setValues(rows);
  sh.setFrozenRows(1);sh.getRange(1,1,1,headers.length).setFontWeight('bold');if(rows.length){sh.getRange(2,8,rows.length,1).setNumberFormat('#,##0');sh.getRange(2,22,rows.length,1).setNumberFormat('#,##0');}
  try{sh.getRange(1,1,rows.length+1,headers.length).createFilter();}catch(_e){}
}
function i72Status_(ss,pairs){var sh=ss.getSheetByName(I72_STATUS)||ss.insertSheet(I72_STATUS);sh.clearContents();sh.getRange(1,1,1,2).setValues([['항목','값']]);if(pairs&&pairs.length)sh.getRange(2,1,pairs.length,2).setValues(pairs);sh.setFrozenRows(1);sh.getRange(1,1,1,2).setFontWeight('bold');sh.autoResizeColumns(1,2);}

function i72FindHeaderRow_(values,required){for(var r=0;r<Math.min(values.length,30);r++){var h=(values[r]||[]).map(i72Text_);var ok=required.every(function(n){return i72Ix_(h,[n])>=0;});if(ok)return r;}return -1;}
function i72Ix_(h,names){var map={};(h||[]).forEach(function(v,i){map[i72Compact_(v)]=i;});for(var j=0;j<names.length;j++){var k=i72Compact_(names[j]);if(Object.prototype.hasOwnProperty.call(map,k))return map[k];}return -1;}
function i72Text_(v){if(v===null||v===undefined)return '';if(Object.prototype.toString.call(v)==='[object Date]')return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');return String(v).trim();}
function i72Compact_(v){return i72Text_(v).toLowerCase().replace(/[\s_\-\/.()\[\]:]+/g,'');}
function i72Num_(v){if(typeof v==='number')return isNaN(v)?0:v;var s=String(v===null||v===undefined?'':v).replace(/,/g,'').replace(/[^0-9.\-]/g,'');var n=Number(s);return isNaN(n)?0:n;}
function i72Date_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');var s=i72Text_(v);var m=s.match(/(20\d{2})[^0-9]?(\d{1,2})[^0-9]?(\d{1,2})/);if(!m)return '';return m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2);}
function i72Days_(a,b){var x=String(a||'').split('-'),y=String(b||'').split('-');if(x.length!==3||y.length!==3)return 99999;return Math.round((Date.UTC(+y[0],+y[1]-1,+y[2])-Date.UTC(+x[0],+x[1]-1,+x[2]))/86400000);}
function i72IsNonCardEvidence_(h){var s=i72Compact_([h.company,h.cardName,h.evidenceType].join(' '));return s.indexOf('비카드')>=0||s.indexOf('현금영수증')>=0||s.indexOf('페이머니')>=0||s.indexOf('머니')>=0;}
function i72IsLotteEvidence_(h){return i72Text_(h.lotteFlag).toUpperCase()==='Y'||/롯데|LOTTE/i.test(i72Text_(h.merchant));}
function i72IsCancel_(h){var s=i72Compact_(h.status);if(!s)return h.amount<0;if(s.indexOf('취소있음')>=0)return false;return s.indexOf('취소')>=0||s.indexOf('환불')>=0||h.amount<0;}
function i72Sig_(sh){var r=sh.getLastRow(),c=sh.getLastColumn();if(r<1||c<1)return r+'x'+c+':EMPTY';var v=sh.getRange(1,1,r,c).getValues();var t=JSON.stringify(v,function(_k,x){if(Object.prototype.toString.call(x)==='[object Date]')return {__date__:x.getTime()};if(typeof x==='number'&&isNaN(x))return {__nan__:true};return x;});var b=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,t,Utilities.Charset.UTF_8);return r+'x'+c+':'+b.map(function(x){var n=(x<0?x+256:x).toString(16);return n.length===1?'0'+n:n;}).join('');}
