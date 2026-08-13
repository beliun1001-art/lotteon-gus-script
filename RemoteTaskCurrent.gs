/**
 * Issue #58 v1.0 read-only cause split for 4 NON_CARD suspect orders.
 * Compares old reused payment, current corrected VAT payment, no payment filter,
 * and physical-card-only matching under the same Issue54 used-before evidence set.
 * Writes only ISSUE58_* sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE58-v1.0-20260813',
  title: 'Issue57 의심 4건 결제수단 재사용 vs current VAT 원인분리',
  enabled: true,
  outputSheet: 'ISSUE58_4건원인분리',
  statusSheet: 'ISSUE58_진단상태'
};
var ISSUE58_CODE_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Code.gs';
var ISSUE58_BOOT_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_24_bootstrap_auto_continue.gs';

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var status = i58Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  i58Write_(status, [
    ['항목','값'],['버전','v1.0-ISSUE58-PAYMENT-SOURCE-CAUSE-SPLIT'],['상태','RUNNING'],['단계','LOAD'],
    ['메시지','의심 4건 old/current/no-payment 원인분리 시작'],['운영시트 변경','0']
  ]);

  var watched = ['부가세_카드매칭검증','부가세_신고자료','카드사용내역_붙여넣기','카드_마스터',
                 'ISSUE54_카드매칭전체PREVIEW','ISSUE55_카드매칭차이진단',
                 'ISSUE56_상태변경12건판정','ISSUE57_차단9건심층진단'];
  var before = {};
  try {
    watched.forEach(function(name) {
      var sh = ss.getSheetByName(name);
      if (sh) before[name] = i58Sig_(sh.getDataRange().getValues());
    });

    var result = i58RunCore_(ss);
    var out = i58Ensure_(ss, LOTTEON_REMOTE_TASK.outputSheet);
    out.clearContents();
    var headers = [
      '쿠팡계정ID','주문번호','주문일','corrected매입금액',
      'Issue54상태','Issue54결제수단','currentVAT결제수단','결제수단동일',
      'OLD_REUSE상태','CURRENT_VAT상태','NO_PAYMENT상태','PHYSICAL_ONLY상태',
      'OLD_REUSE근거','CURRENT_VAT근거','NO_PAYMENT근거','PHYSICAL_ONLY근거',
      'exactPhysical전체','exactPhysical미사용','exactNonCard전체','exactNonCard미사용',
      'physical후보요약','nonCard후보요약','원인분류'
    ];
    out.getRange(1,1,1,headers.length).setValues([headers]);
    if (result.rows.length) out.getRange(2,1,result.rows.length,headers.length).setValues(result.rows);
    out.getRange(1,1,1,headers.length).setFontWeight('bold');
    out.setFrozenRows(1);
    SpreadsheetApp.flush();

    watched.forEach(function(name) {
      var sh = ss.getSheetByName(name);
      if (before[name] && sh) i58Req_(i58Sig_(sh.getDataRange().getValues()) === before[name], name + '이 진단 중 변경되었습니다.');
    });

    var s = result.stats;
    var rows = [
      ['항목','값'],['버전','v1.0-ISSUE58-PAYMENT-SOURCE-CAUSE-SPLIT'],['상태','PASS'],['단계','DONE'],
      ['메시지','의심 4건 old/current/no-payment 원인분리 완료'],['운영시트 변경','0'],
      ['진단대상',s.targets],['canonical증빙행',s.canonicalRows],
      ['OLD_PAYMENT_REUSE_BUG',s.OLD_PAYMENT_REUSE_BUG || 0],
      ['MATCHER_PAYMENT_PRIORITY_BUG',s.MATCHER_PAYMENT_PRIORITY_BUG || 0],
      ['BOTH_PAYMENT_AND_MATCHER',s.BOTH_PAYMENT_AND_MATCHER || 0],
      ['ALLOCATION_CONFLICT',s.ALLOCATION_CONFLICT || 0],
      ['DIAGNOSTIC_FALSE_POSITIVE',s.DIAGNOSTIC_FALSE_POSITIVE || 0],
      ['UNRESOLVED',s.UNRESOLVED || 0],
      ['old/current결제수단다름',s.paymentDifferent || 0],
      ['CURRENT_VAT에서MATCHED',s.currentMatched || 0],
      ['NO_PAYMENT에서MATCHED',s.noPaymentMatched || 0],
      ['PHYSICAL_ONLY에서MATCHED',s.physicalMatched || 0],
      ['exactPhysical미사용후보존재',s.physicalUnusedExists || 0],
      ['운영반영자동승인','NO'],
      ['부가세_카드매칭검증 변경','0'],['부가세_신고자료 변경','0'],
      ['카드사용내역_붙여넣기 변경','0'],['카드_마스터 변경','0']
    ];
    result.summaries.forEach(function(x,i) { rows.push(['확인_'+(i+1),x]); });
    rows.push(['완료시각',new Date().toISOString()]);
    i58Write_(status, rows);
    return {ok:true,done:true,stats:s};
  } catch (e) {
    i58Write_(status, [
      ['항목','값'],['버전','v1.0-ISSUE58-PAYMENT-SOURCE-CAUSE-SPLIT'],['상태','ERROR'],['단계','FAILED'],
      ['메시지','의심 4건 결제수단 원인분리 실패'],['오류',String(e && e.message ? e.message : e)],
      ['운영시트 변경','0'],['완료시각',new Date().toISOString()]
    ]);
    throw e;
  }
}

function runLotteonRemoteTaskContinueRemote_() {
  return {ok:true,done:true,reason:'NO_CONTINUE_REQUIRED'};
}

function i58RunCore_(ss) {
  var code = i58Fetch_(ISSUE58_CODE_URL);
  var boot = i58Fetch_(ISSUE58_BOOT_URL);
  var sid = ss.getId();
  var invocation = [
    ';(function(){',
    'function T(v){return String(v==null?"":v).trim();}',
    'function N(v){return T(v).toLowerCase().replace(/[^0-9a-z가-힣]/g,"");}',
    'function K(a,o){return T(a).toLowerCase()+"|"+N(o);}',
    'function CP(o){var x={};Object.keys(o||{}).forEach(function(k){x[k]=o[k];});return x;}',
    'function CU(u){var x={};Object.keys(u||{}).forEach(function(k){x[k]=true;});return x;}',
    'function ST(r){return T(r&&r.status||"NO_MATCH");}',
    'function EN(h){var e=typeof enrichHistoryFromMaster_v660_==="function"?enrichHistoryFromMaster_v660_(h,master):h;var num=T(e.cardNumber||h.cardNumber);var end4=typeof normalizeVatCardEnd4_v667_==="function"?normalizeVatCardEnd4_v667_(e.cardEnd4||h.cardEnd4,num):T(e.cardEnd4||h.cardEnd4).replace(/\\D/g,"").slice(-4);return {company:T(e.company||h.company),cardName:T(e.cardName||h.cardName),end4:end4,key:T(h.v664CanonicalKey),date:T(h.date),amount:Number(h.v664EffectiveAmount||h.amount||0),nonCard:!!h.nonCard};}',
    'var targets={"beliun1023|17100198549471":true,"beliun1023|4101249972829":true,"beliun1024|3101250601802":true,"beliun1024|31100196211237":true};',
    'var ss=SpreadsheetApp.openById(' + JSON.stringify(sid) + ');',
    'var vat=ss.getSheetByName("부가세_신고자료"),pv=ss.getSheetByName("ISSUE54_카드매칭전체PREVIEW");',
    'if(!vat||vat.getLastRow()<2)throw new Error("부가세_신고자료가 없습니다.");',
    'if(!pv||pv.getLastRow()<2)throw new Error("ISSUE54_카드매칭전체PREVIEW가 없습니다.");',
    'var orders=groupVatDetailByOrder_v660_(vat.getDataRange().getValues()).filter(function(o){return String(o.year)==="2026"&&String(o.half)==="상반기";});',
    'orders.sort(function(a,b){return String(a.orderDate||"").localeCompare(String(b.orderDate||""))||String(a.orderNo||"").localeCompare(String(b.orderNo||""))||Number(a.purchase||0)-Number(b.purchase||0);});',
    'if(orders.length!==1355)throw new Error("current VAT 주문수 불일치: "+orders.length);',
    'var om={};orders.forEach(function(o){om[K(o.account,o.orderNo)]=o;});',
    'var hist=loadVatCardHistory_v660_(ss),master=loadVatCardMaster_v660_(ss),canonical=canonicalizeVatHistory_v664_(hist,master);',
    'if(canonical.length!==1990)throw new Error("canonical 증빙행 불일치: "+canonical.length);',
    'var vals=pv.getDataRange().getValues(),h=vals[0].map(T),ix={a:h.indexOf("쿠팡계정ID"),o:h.indexOf("주문번호"),p:h.indexOf("주문매입금액"),pay:h.indexOf("롯데결제수단"),st:h.indexOf("카드매칭상태"),ck:h.indexOf("canonicalEvidenceKey")};',
    'if(ix.a<0||ix.o<0||ix.p<0||ix.pay<0||ix.st<0||ix.ck<0)throw new Error("Issue54 preview 필수헤더 누락");',
    'if(vals.length-1!==1355)throw new Error("Issue54 preview 주문수 불일치: "+(vals.length-1));',
    'var counts={MATCHED:0,NON_CARD:0,AMBIGUOUS:0,NO_MATCH:0},psum=0;',
    'for(var z=1;z<vals.length;z++){var ssx=T(vals[z][ix.st])||"NO_MATCH";counts[ssx]=(counts[ssx]||0)+1;psum+=Number(vals[z][ix.p]||0);}',
    'if(counts.MATCHED!==808||counts.NON_CARD!==498||counts.AMBIGUOUS!==0||counts.NO_MATCH!==49)throw new Error("Issue54 상태기준 불일치 "+JSON.stringify(counts));',
    'if(Math.round(psum)!==105762969)throw new Error("Issue54 매입합계 불일치: "+Math.round(psum));',
    'var previewMap={},used={};',
    'for(var r=1;r<vals.length;r++){var key=K(vals[r][ix.a],vals[r][ix.o]);var ub=CU(used);previewMap[key]={row:r,oldPayment:T(vals[r][ix.pay]),status:T(vals[r][ix.st]),purchase:Number(vals[r][ix.p]||0),usedBefore:ub};var ckey=T(vals[r][ix.ck]);if(ckey)used[ckey]=true;}',
    'var rows=[],stats={targets:0,canonicalRows:canonical.length,paymentDifferent:0,currentMatched:0,noPaymentMatched:0,physicalMatched:0,physicalUnusedExists:0},summaries=[];',
    'Object.keys(targets).sort().forEach(function(tk){',
      'var o=om[tk],pr=previewMap[tk];if(!o||!pr)throw new Error("대상 주문 누락: "+tk);',
      'if(pr.status!=="NON_CARD")throw new Error("대상 Issue54 상태가 NON_CARD 아님: "+tk+" / "+pr.status);',
      'stats.targets++;var oldPay=pr.oldPayment,currentPay=T(o.lottePayment),same=oldPay===currentPay;if(!same)stats.paymentDifferent++;',
      'var usedBefore=pr.usedBefore;',
      'function RUN(payment,physicalOnly){var oo=CP(o);oo.lottePayment=payment;var hh=physicalOnly?canonical.filter(function(x){return x&&!x.nonCard;}):canonical;return matchVatOrderCardCanonical_v664_(oo,hh,master,CU(usedBefore))||noMatch_v660_("미실행");}',
      'var a=RUN(oldPay,false),b=RUN(currentPay,false),c=RUN("",false),d=RUN("",true);',
      'if(ST(b)==="MATCHED")stats.currentMatched++;if(ST(c)==="MATCHED")stats.noPaymentMatched++;if(ST(d)==="MATCHED")stats.physicalMatched++;',
      'var physical=[],noncard=[];',
      'canonical.forEach(function(x){if(!x||x.cancelRow||x.v664FullyCanceled||!x.lotteEvidence)return;var amt=Number(x.v664EffectiveAmount||x.amount||0);if(amt!==Number(o.purchase||0))return;var dd=typeof daysBetween_v664_==="function"?daysBetween_v664_(o.orderDate,x.date):99999;if(dd<0||dd>7)return;var e=EN(x);e.usedBefore=!!usedBefore[e.key];(x.nonCard?noncard:physical).push(e);});',
      'var pu=physical.filter(function(x){return !x.usedBefore;});var nu=noncard.filter(function(x){return !x.usedBefore;});if(pu.length)stats.physicalUnusedExists++;',
      'function SUM(arr){return arr.map(function(x){return [x.date,x.amount,x.company,x.cardName,x.end4,x.key,(x.usedBefore?"USED":"FREE")].join("/");}).join(" || ");}',
      'var cause="UNRESOLVED";',
      'if(physical.length&&pu.length===0)cause="ALLOCATION_CONFLICT";',
      'else if(ST(a)==="NON_CARD"&&ST(b)==="MATCHED")cause="OLD_PAYMENT_REUSE_BUG";',
      'else if(ST(a)==="NON_CARD"&&ST(b)==="NON_CARD"&&ST(d)==="MATCHED")cause=same?"MATCHER_PAYMENT_PRIORITY_BUG":"BOTH_PAYMENT_AND_MATCHER";',
      'else if(!pu.length&&ST(d)!=="MATCHED")cause="DIAGNOSTIC_FALSE_POSITIVE";',
      'stats[cause]=(stats[cause]||0)+1;',
      'rows.push([o.account,o.orderNo,o.orderDate,Number(o.purchase||0),pr.status,oldPay,currentPay,same?"Y":"N",ST(a),ST(b),ST(c),ST(d),T(a.reason),T(b.reason),T(c.reason),T(d.reason),physical.length,pu.length,noncard.length,nu.length,SUM(physical),SUM(noncard),cause]);',
      'summaries.push(tk+" | old="+oldPay+" | current="+currentPay+" | "+ST(a)+"/"+ST(b)+"/"+ST(c)+"/"+ST(d)+" | physical "+pu.length+"/"+physical.length+" | noncard "+nu.length+"/"+noncard.length+" | "+cause);',
    '});',
    'if(stats.targets!==4)throw new Error("진단대상 불일치: "+stats.targets);',
    'return {rows:rows,stats:stats,summaries:summaries};',
    '})()'
  ].join('\n');
  return eval(code + '\n\n;\n\n' + boot + '\n\n;\n\n' + invocation);
}

function i58Fetch_(url) {
  var res = UrlFetchApp.fetch(url + '?ts=' + new Date().getTime(), {method:'get',muteHttpExceptions:true,followRedirects:true});
  var code = res.getResponseCode(), text = res.getContentText('UTF-8');
  if (code < 200 || code >= 300) throw new Error('원격 코드 로드 실패 HTTP ' + code + ': ' + url + '\n' + text.slice(0,500));
  return text;
}
function i58Ensure_(ss,name) { return ss.getSheetByName(name) || ss.insertSheet(name); }
function i58Write_(sh,rows) { sh.clearContents(); sh.getRange(1,1,rows.length,2).setValues(rows); sh.getRange(1,1,1,2).setFontWeight('bold'); sh.setFrozenRows(1); SpreadsheetApp.flush(); }
function i58Req_(ok,msg) { if (!ok) throw new Error(msg); }
function i58Sig_(v) { var h=2166136261; for(var r=0;r<v.length;r++){for(var c=0;c<v[r].length;c++){var s=String(v[r][c]==null?'':v[r][c]);for(var i=0;i<s.length;i++){h^=s.charCodeAt(i);h+=(h<<1)+(h<<4)+(h<<7)+(h<<8)+(h<<24);}h^=31;}}return String(h>>>0)+'|'+v.length+'|'+(v[0]?v[0].length:0); }
