var LOTTEON_REMOTE_TASK={id:'ISSUE77-v1.0-20260814',title:'NON_CARD 현금영수증 실제증빙 + NO_MATCH22 최종진단',enabled:true,statusSheet:'ISSUE77_실행상태'};
var I77V='v1.0-ISSUE77-TAX-EVIDENCE-FINAL-DIAGNOSTIC';
var I77D='ISSUE77_경정청구증빙진단',I77B='ISSUE77_사업자별증빙요약',I77S='ISSUE77_실행상태';
var I77CORE=['매출데이터_붙여넣기','부가세_신고자료','부가세_카드매칭검증','부가세_기간별','카드사용내역_붙여넣기','카드_마스터'];

function runLotteonRemoteTaskStartRemote_(){return i77run_();}
function runLotteonRemoteTaskContinueRemote_(){return i77run_();}

function i77run_(){
  var ss=SpreadsheetApp.getActive(),before={};
  try{
    i77status_(ss,[['version',I77V],['상태','RUNNING'],['단계','GUARD'],['메시지','NON_CARD 498 실제 현금영수증 증빙 및 NO_MATCH 22 최종 진단 중']]);
    I77CORE.forEach(function(n){var sh=ss.getSheetByName(n);if(!sh)throw new Error('핵심 시트 누락: '+n);before[n]=i77sig_(sh);});
    var s76=ss.getSheetByName('ISSUE76_최종검증상태');if(!s76)throw new Error('ISSUE76_최종검증상태 누락');
    var m76=i77kv_(s76);
    if(i77t_(m76['상태'])!=='PASS'||i77t_(m76['단계'])!=='DONE'||Number(m76['VAT_주문수'])!==1355||Number(m76['MATCHED'])!==835||Number(m76['NON_CARD'])!==498||Number(m76['NO_MATCH'])!==22||Number(m76['핵심시트변경수'])!==0)throw new Error('Issue76 PASS exact guard 실패');

    var vat=i77vat_(ss.getSheetByName('부가세_신고자료'));
    var card=i77card_(ss.getSheetByName('부가세_카드매칭검증'));
    i77guard_(vat,card);
    var cross=i77cross_(vat,card);if(cross.overlap!==1355||cross.vatOnly||cross.cardOnly||cross.purchaseMismatch||cross.businessMismatch)throw new Error('VAT×CARD 1:1 실패 '+JSON.stringify(cross));

    var hist=i77history_(ss.getSheetByName('카드사용내역_붙여넣기'));
    var d73=i77issue73_(ss);
    var diag=[],biz={},nc={confirmed:0,confirmedPurchase:0,confirmedVat:0,unconfirmed:0,unconfirmedPurchase:0,unconfirmedVat:0,other:0,otherPurchase:0,otherVat:0,historyLinked:0,historyUnlinked:0};
    var nm={MULTI_CANDIDATE:0,USED_BY_OTHER:0,AMOUNT_REVIEW:0,BLOCKED_ZERO:0,BLOCKED_CANCELED:0,OTHER:0};
    var nmPurchase={MULTI_CANDIDATE:0,USED_BY_OTHER:0,AMOUNT_REVIEW:0,BLOCKED_ZERO:0,BLOCKED_CANCELED:0,OTHER:0};
    var nmVat={MULTI_CANDIDATE:0,USED_BY_OTHER:0,AMOUNT_REVIEW:0,BLOCKED_ZERO:0,BLOCKED_CANCELED:0,OTHER:0};

    card.rows.forEach(function(c){
      var v=vat.map[c.key]; if(!v)throw new Error('VAT 주문키 누락 '+c.key);
      if(!biz[c.business])biz[c.business]=i77biz_();
      var b=biz[c.business];
      if(c.status==='MATCHED'){
        b.matchedCount++;b.matchedPurchase+=c.purchase;b.matchedVat+=v.purchaseVat;
        return;
      }
      if(c.status==='NON_CARD'){
        var linked=i77linkHistory_(c,hist);
        if(linked.length)nc.historyLinked++;else nc.historyUnlinked++;
        var evidenceText=[c.payment,c.company,c.alias,c.cardName,c.evidence,c.reason,c.source].concat(linked.map(function(h){return [h.company,h.cardName,h.evidence,h.source,h.memo,h.merchant].join(' ');})).join(' ');
        var explicit=i77cashReceiptExplicit_(evidenceText);
        var cls=explicit?'현금영수증_명시확정':'카카오머니_현금영수증미확인';
        if(explicit){
          nc.confirmed++;nc.confirmedPurchase+=c.purchase;nc.confirmedVat+=v.purchaseVat;
          b.cashConfirmedCount++;b.cashConfirmedPurchase+=c.purchase;b.cashConfirmedVat+=v.purchaseVat;
        }else{
          nc.unconfirmed++;nc.unconfirmedPurchase+=c.purchase;nc.unconfirmedVat+=v.purchaseVat;
          b.cashUnconfirmedCount++;b.cashUnconfirmedPurchase+=c.purchase;b.cashUnconfirmedVat+=v.purchaseVat;
        }
        diag.push(['NON_CARD',cls,c.business,c.account,c.orderDate,c.orderNo,c.purchase,v.purchaseVat,c.payment,c.company,c.alias,c.cardName,c.approvalDate,c.approvalNo,c.approvalAmount,c.merchant,c.evidence,c.source,linked.length,explicit?'YES':'NO',i77historySummary_(linked),'','','','','','']);
        return;
      }
      if(c.status==='NO_MATCH'){
        var k=i77c_(c.account)+'|'+i77c_(c.orderNo),d=d73[k];
        if(!d)throw new Error('현재 NO_MATCH의 Issue73 상세 누락: '+k);
        if(d.verdict==='AUTO_SAFE_DATE_WINDOW'||d.verdict==='AUTO_SAFE_SPLIT')throw new Error('회수완료 대상이 NO_MATCH에 잔류: '+k+' '+d.verdict);
        var bucket=Object.prototype.hasOwnProperty.call(nm,d.verdict)?d.verdict:'OTHER';
        nm[bucket]++;nmPurchase[bucket]+=c.purchase;nmVat[bucket]+=v.purchaseVat;
        b.noMatchCount++;b.noMatchPurchase+=c.purchase;b.noMatchVat+=v.purchaseVat;
        diag.push(['NO_MATCH','미확정',c.business,c.account,c.orderDate,c.orderNo,c.purchase,v.purchaseVat,c.payment,c.company,c.alias,c.cardName,c.approvalDate,c.approvalNo,c.approvalAmount,c.merchant,c.evidence,c.source,0,'NO','',d.cause,d.verdict,d.dayLag,d.candidateCount,d.amountDiff,d.summary]);
        return;
      }
      throw new Error('예상 외 카드상태: '+c.status);
    });

    if(nc.confirmed+nc.unconfirmed+nc.other!==498)throw new Error('NON_CARD 분류합계 불일치');
    if(nm.MULTI_CANDIDATE!==11||nm.USED_BY_OTHER!==1||nm.AMOUNT_REVIEW!==8||nm.BLOCKED_ZERO!==1||nm.BLOCKED_CANCELED!==1||nm.OTHER!==0)throw new Error('NO_MATCH22 Issue73 잔여구조 불일치 '+JSON.stringify(nm));
    if(diag.length!==520)throw new Error('진단 행수 불일치 '+diag.length+' (기대 520)');

    i77writeDiag_(ss,diag);
    var bizRows=i77businessRows_(biz);i77writeBiz_(ss,bizRows);

    var changed=[];I77CORE.forEach(function(n){var sh=ss.getSheetByName(n);if(!sh||i77sig_(sh)!==before[n])changed.push(n);});if(changed.length)throw new Error('READ-ONLY 위반: '+changed.join(', '));

    var st=[
      ['version',I77V],['상태','PASS'],['단계','DONE'],['메시지','NON_CARD 498 실제 현금영수증 명시 여부 및 잔여 NO_MATCH 22 최종 경정청구 진단 완료'],
      ['운영주문',1355],['MATCHED',835],['NON_CARD',498],['NO_MATCH',22],['매입합계',105762969],
      ['NONCARD_현금영수증명시확정',nc.confirmed],['NONCARD_현금영수증명시확정_매입',nc.confirmedPurchase],['NONCARD_현금영수증명시확정_스크립트매입부가세',nc.confirmedVat],
      ['NONCARD_카카오머니_현금영수증미확인',nc.unconfirmed],['NONCARD_카카오머니_미확인_매입',nc.unconfirmedPurchase],['NONCARD_카카오머니_미확인_스크립트매입부가세',nc.unconfirmedVat],
      ['NONCARD_비카드기타',nc.other],['NONCARD_증빙원본연결',nc.historyLinked],['NONCARD_증빙원본미연결',nc.historyUnlinked],
      ['NOMATCH_MULTI_CANDIDATE',nm.MULTI_CANDIDATE],['NOMATCH_MULTI_CANDIDATE_매입',nmPurchase.MULTI_CANDIDATE],['NOMATCH_MULTI_CANDIDATE_스크립트매입부가세',nmVat.MULTI_CANDIDATE],
      ['NOMATCH_USED_BY_OTHER',nm.USED_BY_OTHER],['NOMATCH_USED_BY_OTHER_매입',nmPurchase.USED_BY_OTHER],['NOMATCH_USED_BY_OTHER_스크립트매입부가세',nmVat.USED_BY_OTHER],
      ['NOMATCH_AMOUNT_REVIEW',nm.AMOUNT_REVIEW],['NOMATCH_AMOUNT_REVIEW_매입',nmPurchase.AMOUNT_REVIEW],['NOMATCH_AMOUNT_REVIEW_스크립트매입부가세',nmVat.AMOUNT_REVIEW],
      ['NOMATCH_BLOCKED_ZERO',nm.BLOCKED_ZERO],['NOMATCH_BLOCKED_CANCELED',nm.BLOCKED_CANCELED],
      ['사업자수',bizRows.length],['진단행수',diag.length],['핵심시트변경수',0],['오류',''],['완료시각',new Date().toISOString()]
    ];
    i77status_(ss,st);try{ss.toast('Issue77 PASS: 현금영수증 명시확정 '+nc.confirmed+' / 미확인 '+nc.unconfirmed+' / NO_MATCH 22 최종분류','LOTTEON',10);}catch(_e){}
    return{ok:true,done:true,version:I77V,cashReceiptConfirmed:nc.confirmed,cashReceiptUnconfirmed:nc.unconfirmed,noMatch:22};
  }catch(e){
    var msg=String(e&&e.message?e.message:e);
    try{i77status_(ss,[['version',I77V],['상태','ERROR'],['단계','FAILED'],['메시지','Issue77 증빙 진단 실패'],['오류',msg],['완료시각',new Date().toISOString()]]);}catch(_e2){}
    throw e;
  }
}

function i77vat_(sh){
  var v=sh.getDataRange().getValues(),h=v[0]||[],x={year:i77ix_(h,['신고연도']),half:i77ix_(h,['반기']),date:i77ix_(h,['날짜','주문일']),account:i77ix_(h,['쿠팡계정ID']),business:i77ix_(h,['사업자등록번호']),order:i77ix_(h,['주문번호']),purchase:i77ix_(h,['매입금액']),pv:i77ix_(h,['매입부가세'])};
  Object.keys(x).forEach(function(k){if(x[k]<0)throw new Error('VAT header 누락 '+k);});
  var m={},detail=0,purchase=0,pv=0;
  for(var r=1;r<v.length;r++){var z=v[r];if(i77t_(z[x.year])!=='2026'||i77t_(z[x.half])!=='상반기')continue;var no=i77t_(z[x.order]);if(!no)continue;detail++;var key=[i77t_(z[x.year]),i77t_(z[x.half]),i77t_(z[x.business]),i77t_(z[x.account]),no].join('|');if(!m[key])m[key]={business:i77t_(z[x.business]),account:i77t_(z[x.account]),orderNo:no,orderDate:i77d_(z[x.date]),purchase:0,purchaseVat:0};m[key].purchase+=i77n_(z[x.purchase]);m[key].purchaseVat+=i77n_(z[x.pv]);purchase+=i77n_(z[x.purchase]);pv+=i77n_(z[x.pv]);}
  return{detailRows:detail,orders:Object.keys(m).length,purchase:purchase,purchaseVat:pv,map:m};
}
function i77card_(sh){
  var v=sh.getDataRange().getValues(),hr=i77header_(v,['주문번호','카드매칭상태']);if(hr<0)throw new Error('카드검증 header 탐지 실패');var h=v[hr],x={year:i77ix_(h,['신고연도']),half:i77ix_(h,['반기']),date:i77ix_(h,['주문일']),business:i77ix_(h,['사업자등록번호']),account:i77ix_(h,['쿠팡계정ID']),order:i77ix_(h,['주문번호']),payment:i77ix_(h,['롯데결제수단']),purchase:i77ix_(h,['주문매입금액']),company:i77ix_(h,['구매카드사']),alias:i77ix_(h,['구매카드별칭']),name:i77ix_(h,['구매카드명']),adate:i77ix_(h,['승인일']),approval:i77ix_(h,['승인번호']),aamount:i77ix_(h,['승인금액']),status:i77ix_(h,['카드매칭상태']),reason:i77ix_(h,['카드매칭근거']),merchant:i77ix_(h,['가맹점명']),evidence:i77ix_(h,['증빙유형']),source:i77ix_(h,['원본파일'])};
  ['year','half','date','business','account','order','purchase','status'].forEach(function(k){if(x[k]<0)throw new Error('카드 header 누락 '+k);});
  var rows=[],stats={},purchase=0,keys={};
  for(var r=hr+1;r<v.length;r++){var z=v[r],no=i77t_(z[x.order]);if(!no)continue;var key=[i77t_(z[x.year]),i77t_(z[x.half]),i77t_(z[x.business]),i77t_(z[x.account]),no].join('|');if(keys[key])throw new Error('카드 주문키 중복 '+key);keys[key]=1;var st=i77t_(z[x.status]).toUpperCase(),p=i77n_(z[x.purchase]);stats[st]=(stats[st]||0)+1;purchase+=p;rows.push({key:key,business:i77t_(z[x.business]),account:i77t_(z[x.account]),orderDate:i77d_(z[x.date]),orderNo:no,purchase:p,status:st,payment:x.payment>=0?i77t_(z[x.payment]):'',company:x.company>=0?i77t_(z[x.company]):'',alias:x.alias>=0?i77t_(z[x.alias]):'',cardName:x.name>=0?i77t_(z[x.name]):'',approvalDate:x.adate>=0?i77d_(z[x.adate]):'',approvalNo:x.approval>=0?i77t_(z[x.approval]):'',approvalAmount:x.aamount>=0?i77n_(z[x.aamount]):0,reason:x.reason>=0?i77t_(z[x.reason]):'',merchant:x.merchant>=0?i77t_(z[x.merchant]):'',evidence:x.evidence>=0?i77t_(z[x.evidence]):'',source:x.source>=0?i77t_(z[x.source]):''});}
  return{rows:rows,stats:stats,purchase:purchase,map:(function(){var m={};rows.forEach(function(o){m[o.key]=o;});return m;})()};
}
function i77guard_(vat,card){
  if(vat.detailRows!==2752||vat.orders!==1355||vat.purchase!==105762969||vat.purchaseVat!==9614786)throw new Error('VAT exact guard '+JSON.stringify({detail:vat.detailRows,orders:vat.orders,purchase:vat.purchase,pv:vat.purchaseVat}));
  if(card.rows.length!==1355||(card.stats.MATCHED||0)!==835||(card.stats.NON_CARD||0)!==498||(card.stats.AMBIGUOUS||0)!==0||(card.stats.NO_MATCH||0)!==22||card.purchase!==105762969)throw new Error('카드 exact guard '+JSON.stringify(card.stats));
}
function i77cross_(vat,card){var x={overlap:0,vatOnly:0,cardOnly:0,purchaseMismatch:0,businessMismatch:0},seen={};Object.keys(vat.map).forEach(function(k){var a=vat.map[k],b=card.map[k];if(!b){x.vatOnly++;return;}x.overlap++;seen[k]=1;if(a.purchase!==b.purchase)x.purchaseMismatch++;if(a.business!==b.business)x.businessMismatch++;});Object.keys(card.map).forEach(function(k){if(!seen[k])x.cardOnly++;});return x;}

function i77history_(sh){
  var v=sh.getDataRange().getValues(),hr=i77header_(v,['카드사','승인금액']);if(hr<0)hr=0;var h=v[hr],x={company:i77ix_(h,['카드사']),name:i77ix_(h,['카드명']),date:i77ix_(h,['승인일','이용일','거래일']),merchant:i77ix_(h,['가맹점명','이용가맹점']),amount:i77ix_(h,['승인금액','이용금액','거래금액']),approval:i77ix_(h,['승인번호']),evidence:i77ix_(h,['증빙유형']),source:i77ix_(h,['원본파일']),memo:i77ix_(h,['메모'])};if(x.date<0||x.amount<0)throw new Error('카드사용내역 필수 header 누락');
  var a=[];for(var r=hr+1;r<v.length;r++){var z=v[r],date=i77d_(z[x.date]),amount=i77n_(z[x.amount]);if(!date&&!amount)continue;a.push({row:r+1,company:x.company>=0?i77t_(z[x.company]):'',cardName:x.name>=0?i77t_(z[x.name]):'',date:date,merchant:x.merchant>=0?i77t_(z[x.merchant]):'',amount:amount,approvalNo:x.approval>=0?i77t_(z[x.approval]):'',evidence:x.evidence>=0?i77t_(z[x.evidence]):'',source:x.source>=0?i77t_(z[x.source]):'',memo:x.memo>=0?i77t_(z[x.memo]):''});}return a;
}
function i77linkHistory_(c,h){
  var app=i77c_(c.approvalNo),source=i77c_(c.source),merchant=i77c_(c.merchant),out=[];
  if(app){out=h.filter(function(q){return i77c_(q.approvalNo)===app;});if(out.length)return out;}
  out=h.filter(function(q){if(c.approvalDate&&q.date!==c.approvalDate)return false;if(c.approvalAmount&&q.amount!==c.approvalAmount)return false;var s=source&&i77c_(q.source)===source,m=merchant&&i77c_(q.merchant)===merchant;return s||m;});
  if(out.length)return out;
  return h.filter(function(q){return c.approvalDate&&q.date===c.approvalDate&&c.approvalAmount&&q.amount===c.approvalAmount&&/카카오|머니|비카드/.test(i77c_([q.company,q.cardName,q.evidence].join(' ')));});
}
function i77cashReceiptExplicit_(s){var c=i77c_(s);return c.indexOf('현금영수증')>=0||c.indexOf('cashreceipt')>=0;}
function i77historySummary_(a){return a.slice(0,5).map(function(q){return [q.row,q.date,q.amount,q.company,q.cardName,q.evidence,q.source,q.memo].join('|');}).join(' / ');}

function i77issue73_(ss){
  var sh=ss.getSheetByName('ISSUE73_NOMATCH회수진단');if(!sh)throw new Error('ISSUE73_NOMATCH회수진단 누락');var v=sh.getDataRange().getValues(),h=v[0]||[],x={account:i77ix_(h,['쿠팡계정ID']),order:i77ix_(h,['주문번호']),cause:i77ix_(h,['Issue72원인']),verdict:i77ix_(h,['판정']),lag:i77ix_(h,['dayLag']),count:i77ix_(h,['후보수']),diff:i77ix_(h,['금액차이']),summary:i77ix_(h,['후보요약'])};Object.keys(x).forEach(function(k){if(x[k]<0)throw new Error('Issue73 header 누락 '+k);});var m={};for(var r=1;r<v.length;r++){var no=i77t_(v[r][x.order]);if(!no)continue;var key=i77c_(v[r][x.account])+'|'+i77c_(no);if(m[key])throw new Error('Issue73 주문키 중복 '+key);m[key]={cause:i77t_(v[r][x.cause]),verdict:i77t_(v[r][x.verdict]),dayLag:v[r][x.lag],candidateCount:i77n_(v[r][x.count]),amountDiff:i77n_(v[r][x.diff]),summary:i77t_(v[r][x.summary])};}if(Object.keys(m).length!==49)throw new Error('Issue73 상세행수 '+Object.keys(m).length+' (기대49)');return m;
}

function i77biz_(){return{matchedCount:0,matchedPurchase:0,matchedVat:0,cashConfirmedCount:0,cashConfirmedPurchase:0,cashConfirmedVat:0,cashUnconfirmedCount:0,cashUnconfirmedPurchase:0,cashUnconfirmedVat:0,noMatchCount:0,noMatchPurchase:0,noMatchVat:0};}
function i77businessRows_(m){return Object.keys(m).sort().map(function(k){var b=m[k],total=b.matchedCount+b.cashConfirmedCount+b.cashUnconfirmedCount+b.noMatchCount;return[k,total,b.matchedCount,b.matchedPurchase,b.matchedVat,b.cashConfirmedCount,b.cashConfirmedPurchase,b.cashConfirmedVat,b.cashUnconfirmedCount,b.cashUnconfirmedPurchase,b.cashUnconfirmedVat,b.noMatchCount,b.noMatchPurchase,b.noMatchVat,b.matchedPurchase+b.cashConfirmedPurchase+b.cashUnconfirmedPurchase+b.noMatchPurchase,b.matchedVat+b.cashConfirmedVat+b.cashUnconfirmedVat+b.noMatchVat];});}

function i77writeDiag_(ss,rows){var sh=ss.getSheetByName(I77D)||ss.insertSheet(I77D);if(sh.getFilter())sh.getFilter().remove();sh.clear();var h=['구분','경정청구증빙분류','사업자등록번호','쿠팡계정ID','주문일','주문번호','주문매입금액','스크립트매입부가세','롯데결제수단','구매카드사','구매카드별칭','구매카드명','승인일','승인번호','승인금액','가맹점명','증빙유형','원본파일','원본증빙연결수','현금영수증명시','원본증빙요약','Issue72원인','Issue73판정','dayLag','후보수','금액차이','후보요약'];sh.getRange(1,1,1,h.length).setValues([h]);if(rows.length)sh.getRange(2,1,rows.length,h.length).setValues(rows);sh.setFrozenRows(1);sh.getRange(1,1,1,h.length).setFontWeight('bold');if(rows.length){sh.getRange(2,7,rows.length,2).setNumberFormat('#,##0');sh.getRange(2,15,rows.length,1).setNumberFormat('#,##0');sh.getRange(2,25,rows.length,2).setNumberFormat('#,##0');}try{sh.getRange(1,1,rows.length+1,h.length).createFilter();}catch(_e){}}
function i77writeBiz_(ss,rows){var sh=ss.getSheetByName(I77B)||ss.insertSheet(I77B);sh.clear();var h=['사업자등록번호','총주문수','카드매칭확정건수','카드매칭확정_매입','카드매칭확정_스크립트매입부가세','현금영수증명시확정건수','현금영수증명시확정_매입','현금영수증명시확정_스크립트매입부가세','카카오머니_현금영수증미확인건수','카카오머니_미확인_매입','카카오머니_미확인_스크립트매입부가세','NO_MATCH건수','NO_MATCH_매입','NO_MATCH_스크립트매입부가세','총매입','총스크립트매입부가세'];sh.getRange(1,1,1,h.length).setValues([h]);if(rows.length)sh.getRange(2,1,rows.length,h.length).setValues(rows);sh.setFrozenRows(1);sh.getRange(1,1,1,h.length).setFontWeight('bold');if(rows.length)sh.getRange(2,3,rows.length,h.length-2).setNumberFormat('#,##0');}
function i77status_(ss,p){var sh=ss.getSheetByName(I77S)||ss.insertSheet(I77S);sh.clearContents();sh.getRange(1,1,1,2).setValues([['항목','값']]);if(p&&p.length)sh.getRange(2,1,p.length,2).setValues(p);sh.setFrozenRows(1);sh.getRange(1,1,1,2).setFontWeight('bold');sh.autoResizeColumns(1,2);}

function i77kv_(sh){var v=sh.getDataRange().getValues(),m={};for(var r=1;r<v.length;r++)m[i77t_(v[r][0])]=v[r][1];return m;}
function i77header_(v,req){for(var r=0;r<Math.min(v.length,30);r++){var ok=true;for(var i=0;i<req.length;i++)if(i77ix_(v[r],[req[i]])<0)ok=false;if(ok)return r;}return -1;}
function i77ix_(h,n){var m={};(h||[]).forEach(function(v,i){m[i77c_(v)]=i;});for(var j=0;j<n.length;j++){var k=i77c_(n[j]);if(Object.prototype.hasOwnProperty.call(m,k))return m[k];}return -1;}
function i77t_(v){if(v===null||v===undefined)return'';if(Object.prototype.toString.call(v)==='[object Date]')return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');return String(v).trim();}
function i77c_(v){return i77t_(v).toLowerCase().replace(/[\s_\-\/.()\[\]:]+/g,'');}
function i77n_(v){if(typeof v==='number')return isNaN(v)?0:Math.round(v);var n=Number(String(v==null?'0':v).replace(/[원,\s]/g,''));return isNaN(n)?0:Math.round(n);}
function i77d_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');var s=i77t_(v),m=s.match(/(20\d{2})[^0-9]?(\d{1,2})[^0-9]?(\d{1,2})/);return m?m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2):s;}
function i77sig_(sh){var r=sh.getLastRow(),c=sh.getLastColumn();if(r<1||c<1)return r+'x'+c+':EMPTY';var v=sh.getRange(1,1,r,c).getValues(),s=JSON.stringify(v,function(_k,x){if(Object.prototype.toString.call(x)==='[object Date]')return{__date__:x.getTime()};if(typeof x==='number'&&isNaN(x))return{__nan__:true};return x;});var b=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8);return r+'x'+c+':'+b.map(function(x){var n=(x<0?x+256:x).toString(16);return n.length===1?'0'+n:n;}).join('');}
