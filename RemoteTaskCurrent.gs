var LOTTEON_REMOTE_TASK={id:'ISSUE74-v1.0-20260814',title:'NO_MATCH 안전회수 matcher 규칙 PREVIEW',enabled:true,statusSheet:'ISSUE74_실행상태'};
var I74_VERSION='v1.0-ISSUE74-SAFE-RECOVERY-MATCHER-PREVIEW';
var I74_PREVIEW='ISSUE74_카드회수PREVIEW',I74_STATUS='ISSUE74_실행상태';
var I74_CORE=['매출데이터_붙여넣기','부가세_신고자료','부가세_카드매칭검증','부가세_기간별','카드사용내역_붙여넣기','카드_마스터'];

function runLotteonRemoteTaskStartRemote_(){return i74Run_();}
function runLotteonRemoteTaskContinueRemote_(){return i74Run_();}

function i74Run_(){
  var ss=SpreadsheetApp.getActive(),before={};
  try{
    i74Status_(ss,[['version',I74_VERSION],['상태','RUNNING'],['단계','GUARD'],['메시지','Issue73 AUTO_SAFE 27건 순차할당 preview 준비 중']]);
    I74_CORE.forEach(function(n){var s=ss.getSheetByName(n);if(!s)throw new Error('핵심 시트 누락: '+n);before[n]=i74Sig_(s);});

    var eligible=i74Eligible73_(ss);
    var card=i74ReadCard_(ss.getSheetByName('부가세_카드매칭검증'));
    i74GuardCard_(card);
    var raw=i74ReadHist_(ss.getSheetByName('카드사용내역_붙여넣기'));
    var canon=i74Canon_(raw);
    var used=i74UsedSet_(card.rows);
    var preview=card.values.map(function(r){return r.slice();});

    var recoveredDate=0,recoveredDatePurchase=0,recoveredSplit=0,recoveredSplitPurchase=0;
    var collisionDate=0,collisionSplit=0,splitNonCardBlocked=0,eligibleSeen=0;
    var recoveryRows=[],recoveredKeys={};

    var noMatch=card.rows.filter(function(o){return o.status==='NO_MATCH';}).slice().sort(function(a,b){
      return String(a.orderDate).localeCompare(String(b.orderDate))||String(a.orderNo).localeCompare(String(b.orderNo))||Number(a.purchase)-Number(b.purchase);
    });

    noMatch.forEach(function(o){
      var key=i74C_(o.account)+'|'+i74C_(o.orderNo),kind=eligible[key]||'';
      if(!kind)return;
      eligibleSeen++;
      if(kind==='AUTO_SAFE_DATE_WINDOW'){
        var d=i74RecoverDate_(o,canon,used);
        if(!d.ok){collisionDate++;recoveryRows.push([o.business,o.account,o.orderDate,o.orderNo,o.purchase,kind,'REVIEW_'+d.reason,'','','',d.summary]);return;}
        i74ApplySingle_(preview[o.valueRow],card.ix,o,d.candidate);
        used[d.candidate.key]=true;recoveredKeys[key]=true;
        recoveredDate++;recoveredDatePurchase+=o.purchase;
        recoveryRows.push([o.business,o.account,o.orderDate,o.orderNo,o.purchase,kind,'RECOVERED_DATE',d.candidate.date,d.candidate.company,d.candidate.identity,d.summary]);
      }else if(kind==='AUTO_SAFE_SPLIT'){
        var s=i74RecoverSplit_(o,canon,used);
        if(!s.ok){if(s.reason==='NONCARD_PAIR')splitNonCardBlocked++;else collisionSplit++;recoveryRows.push([o.business,o.account,o.orderDate,o.orderNo,o.purchase,kind,'REVIEW_'+s.reason,'','','',s.summary]);return;}
        i74ApplySplit_(preview[o.valueRow],card.ix,o,s.pair);
        used[s.pair[0].key]=true;used[s.pair[1].key]=true;recoveredKeys[key]=true;
        recoveredSplit++;recoveredSplitPurchase+=o.purchase;
        recoveryRows.push([o.business,o.account,o.orderDate,o.orderNo,o.purchase,kind,'RECOVERED_SPLIT',s.pair[0].date,s.pair[0].company,s.pair[0].identity,s.summary]);
      }
    });

    if(eligibleSeen!==27)throw new Error('Issue73 AUTO_SAFE 대상 연결 실패: '+eligibleSeen+' / 27');
    var stats=i74StatsFromMatrix_(preview,card.hr,card.ix);
    var changed=i74ChangedRows_(card.values,preview,card.hr,card.ix,eligible,recoveredKeys);
    if(changed.unexpected)throw new Error('AUTO_SAFE 외 preview 변경 감지: '+changed.unexpected);
    if(stats.total!==1355||stats.purchase!==105762969)throw new Error('preview 총계 guard 실패: '+JSON.stringify(stats));
    if(stats.NON_CARD!==498||stats.AMBIGUOUS!==0)throw new Error('preview 보호 상태 집계 변경: '+JSON.stringify(stats));
    if(stats.MATCHED+stats.NO_MATCH!==857)throw new Error('preview MATCHED+NO_MATCH 합계 불일치');

    i74WritePreview_(ss,preview);
    i74WriteRecoveryDetail_(ss,recoveryRows);

    var coreChanged=[];
    I74_CORE.forEach(function(n){var s=ss.getSheetByName(n);if(!s||i74Sig_(s)!==before[n])coreChanged.push(n);});
    if(coreChanged.length)throw new Error('보호시트 변경 감지: '+coreChanged.join(', '));

    var remaining=49-recoveredDate-recoveredSplit;
    var st=[
      ['version',I74_VERSION],['상태','PASS'],['단계','DONE'],['메시지','Issue73 AUTO_SAFE 후보의 상호 증빙충돌 포함 순차할당 PREVIEW 완료'],
      ['운영주문',1355],['운영_MATCHED',808],['운영_NON_CARD',498],['운영_AMBIGUOUS',0],['운영_NO_MATCH',49],
      ['Issue73_DATE_SAFE',26],['Issue73_SPLIT_SAFE',1],['Issue73_SAFE합계',27],
      ['PREVIEW_DATE_RECOVERED',recoveredDate],['PREVIEW_DATE_RECOVERED_매입',recoveredDatePurchase],
      ['PREVIEW_SPLIT_RECOVERED',recoveredSplit],['PREVIEW_SPLIT_RECOVERED_매입',recoveredSplitPurchase],
      ['PREVIEW_총회수',recoveredDate+recoveredSplit],['PREVIEW_총회수_매입',recoveredDatePurchase+recoveredSplitPurchase],
      ['DATE_상호충돌_REVIEW',collisionDate],['SPLIT_상호충돌_REVIEW',collisionSplit],['SPLIT_NONCARD_BLOCKED',splitNonCardBlocked],
      ['PREVIEW_MATCHED',stats.MATCHED],['PREVIEW_NON_CARD',stats.NON_CARD],['PREVIEW_AMBIGUOUS',stats.AMBIGUOUS],['PREVIEW_NO_MATCH',stats.NO_MATCH],
      ['PREVIEW_매입합계',stats.purchase],['PREVIEW_잔여NO_MATCH',remaining],
      ['AUTO_SAFE외변경행',changed.unexpected],['실제변경행',changed.changed],['기존_MATCHED_NONCARD_변경행',changed.protectedChanged],
      ['보호시트변경수',0],['오류',''],['완료시각',new Date().toISOString()]
    ];
    i74Status_(ss,st);
    try{ss.toast('Issue74 PASS: 안전회수 '+(recoveredDate+recoveredSplit)+'건 / 잔여 NO_MATCH '+remaining+'건','LOTTEON',10);}catch(_e){}
    return {ok:true,done:true,version:I74_VERSION,recovered:recoveredDate+recoveredSplit,noMatch:remaining};
  }catch(e){
    var m=String(e&&e.message?e.message:e);
    try{i74Status_(ss,[['version',I74_VERSION],['상태','ERROR'],['단계','FAILED'],['메시지','Issue74 preview 실패'],['오류',m],['완료시각',new Date().toISOString()]]);}catch(_e2){}
    throw e;
  }
}

function i74Eligible73_(ss){
  var st=ss.getSheetByName('ISSUE73_실행상태');if(!st)throw new Error('ISSUE73_실행상태 누락');
  var sv=st.getDataRange().getValues(),m={};for(var r=1;r<sv.length;r++)m[i74T_(sv[r][0])]=sv[r][1];
  if(i74T_(m['상태'])!=='PASS'||i74T_(m['단계'])!=='DONE'||Number(m['NO_MATCH'])!==49||Number(m['AUTO_SAFE_DATE_WINDOW'])!==26||Number(m['AUTO_SAFE_SPLIT'])!==1||Number(m['보호시트변경수'])!==0)throw new Error('Issue73 PASS exact guard 실패');
  var sh=ss.getSheetByName('ISSUE73_NOMATCH회수진단');if(!sh)throw new Error('ISSUE73_NOMATCH회수진단 누락');
  var v=sh.getDataRange().getValues(),h=v[0],a=i74Ix_(h,['쿠팡계정ID']),o=i74Ix_(h,['주문번호']),x=i74Ix_(h,['판정','verdict','회수판정']);
  if(x<0)x=i74FindHeaderContains_(h,'판정');
  if(a<0||o<0||x<0)throw new Error('Issue73 상세 header 탐지 실패');
  var out={},dc=0,sc=0;
  for(var i=1;i<v.length;i++){
    var verdict=i74T_(v[i][x]),no=i74T_(v[i][o]);if(!no)continue;
    if(verdict==='AUTO_SAFE_DATE_WINDOW'||verdict==='AUTO_SAFE_SPLIT'){
      var k=i74C_(v[i][a])+'|'+i74C_(no);if(out[k])throw new Error('Issue73 SAFE 주문키 중복: '+k);out[k]=verdict;
      if(verdict==='AUTO_SAFE_DATE_WINDOW')dc++;else sc++;
    }
  }
  if(dc!==26||sc!==1)throw new Error('Issue73 SAFE 상세 합계 불일치: '+dc+'/'+sc);
  return out;
}

function i74ReadCard_(sh){
  var values=sh.getDataRange().getValues(),hr=i74Header_(values,['주문번호','카드매칭상태']);if(hr<0)throw new Error('카드검증 header 탐지 실패');
  var h=values[hr],ix={
    date:i74Ix_(h,['주문일']),business:i74Ix_(h,['사업자등록번호']),account:i74Ix_(h,['쿠팡계정ID']),order:i74Ix_(h,['주문번호']),purchase:i74Ix_(h,['주문매입금액']),payment:i74Ix_(h,['롯데결제수단']),
    company:i74Ix_(h,['구매카드사']),alias:i74Ix_(h,['구매카드별칭']),cardName:i74Ix_(h,['구매카드명']),number:i74Ix_(h,['카드번호']),end4:i74Ix_(h,['카드번호끝4']),
    adate:i74Ix_(h,['승인일']),atime:i74Ix_(h,['승인시각']),approval:i74Ix_(h,['승인번호']),aamount:i74Ix_(h,['승인금액']),status:i74Ix_(h,['카드매칭상태']),reason:i74Ix_(h,['카드매칭근거']),candidate:i74Ix_(h,['후보수']),
    merchant:i74Ix_(h,['가맹점명']),merchantOrder:i74Ix_(h,['가맹점주문번호']),evidence:i74Ix_(h,['증빙유형']),cancelMemo:i74Ix_(h,['취소/부분취소메모']),source:i74Ix_(h,['원본파일']),summary:i74Ix_(h,['후보요약'])
  };
  ['date','business','account','order','purchase','status'].forEach(function(k){if(ix[k]<0)throw new Error('카드검증 필수 header 누락: '+k);});
  var rows=[],stats={},purchase=0,keys={};
  for(var r=hr+1;r<values.length;r++){
    var z=values[r],no=i74T_(z[ix.order]);if(!no)continue;
    var account=i74T_(z[ix.account]),key=i74C_(account)+'|'+i74C_(no);if(keys[key])throw new Error('카드검증 주문키 중복: '+key);keys[key]=1;
    var s=i74T_(z[ix.status]).toUpperCase(),p=i74N_(z[ix.purchase]);stats[s]=(stats[s]||0)+1;purchase+=p;
    rows.push({valueRow:r,business:i74T_(z[ix.business]),account:account,orderDate:i74D_(z[ix.date]),orderNo:no,purchase:p,payment:ix.payment>=0?i74T_(z[ix.payment]):'',status:s,
      company:ix.company>=0?i74T_(z[ix.company]):'',approvalNo:ix.approval>=0?i74T_(z[ix.approval]):'',approvalDate:ix.adate>=0?i74D_(z[ix.adate]):'',approvalAmount:ix.aamount>=0?i74N_(z[ix.aamount]):0,merchant:ix.merchant>=0?i74T_(z[ix.merchant]):'',source:ix.source>=0?i74T_(z[ix.source]):''});
  }
  return {values:values,hr:hr,h:h,ix:ix,rows:rows,stats:stats,purchase:Math.round(purchase)};
}
function i74GuardCard_(x){if(x.rows.length!==1355||(x.stats.MATCHED||0)!==808||(x.stats.NON_CARD||0)!==498||(x.stats.AMBIGUOUS||0)!==0||(x.stats.NO_MATCH||0)!==49||x.purchase!==105762969)throw new Error('운영 카드검증 exact guard 실패');}

function i74ReadHist_(sh){
  var v=sh.getDataRange().getValues(),hr=i74Header_(v,['카드사','승인금액']);if(hr<0)hr=0;var h=v[hr];
  var x={company:i74Ix_(h,['카드사']),name:i74Ix_(h,['카드명']),number:i74Ix_(h,['카드번호']),end4:i74Ix_(h,['카드번호끝4']),date:i74Ix_(h,['승인일','이용일','거래일']),time:i74Ix_(h,['승인시각','이용시각','거래시각']),merchant:i74Ix_(h,['가맹점명','이용가맹점']),amount:i74Ix_(h,['승인금액','이용금액','거래금액']),approval:i74Ix_(h,['승인번호']),status:i74Ix_(h,['승인상태','승인/취소구분','상태']),cancelDate:i74Ix_(h,['취소일']),cancelAmount:i74Ix_(h,['취소금액']),orderNo:i74Ix_(h,['가맹점주문번호','주문번호']),evidence:i74Ix_(h,['증빙유형']),lotte:i74Ix_(h,['롯데계열여부']),source:i74Ix_(h,['원본파일']),memo:i74Ix_(h,['메모'])};
  if(x.date<0||x.amount<0)throw new Error('카드사용내역 필수 header 누락');
  var out=[];
  for(var r=hr+1;r<v.length;r++){
    var z=v[r],date=i74D_(z[x.date]),amt=i74N_(z[x.amount]);if(!date&&!amt)continue;
    var q={row:r+1,company:x.company>=0?i74T_(z[x.company]):'',cardName:x.name>=0?i74T_(z[x.name]):'',cardNumber:x.number>=0?i74T_(z[x.number]):'',end4:x.end4>=0?i74T_(z[x.end4]):'',date:date,time:x.time>=0?i74T_(z[x.time]):'',merchant:x.merchant>=0?i74T_(z[x.merchant]):'',amount:amt,approvalNo:x.approval>=0?i74T_(z[x.approval]):'',status:x.status>=0?i74T_(z[x.status]):'',cancelDate:x.cancelDate>=0?i74D_(z[x.cancelDate]):'',cancelAmount:x.cancelAmount>=0?i74N_(z[x.cancelAmount]):0,merchantOrderNo:x.orderNo>=0?i74T_(z[x.orderNo]):'',evidence:x.evidence>=0?i74T_(z[x.evidence]):'',lotte:x.lotte>=0?i74T_(z[x.lotte]):'',source:x.source>=0?i74T_(z[x.source]):'',memo:x.memo>=0?i74T_(z[x.memo]):''};
    q.nonCard=/비카드|현금영수증|페이머니|머니/.test(i74C_([q.company,q.cardName,q.evidence].join(' ')));q.cancel=i74Cancel_(q);q.lotteEvidence=i74T_(q.lotte).toUpperCase()==='Y'||/롯데|LOTTE/i.test(q.merchant);out.push(q);
  }
  return out;
}
function i74Canon_(raw){
  var groups={},out=[];
  raw.forEach(function(q){var app=i74C_(q.approvalNo),co=i74C_(q.company);if(!q.nonCard&&app&&co){var k='APP|'+co+'|'+app;if(!groups[k])groups[k]=[];groups[k].push(q);}else{var x=i74Single_(q);if(x)out.push(x);}});
  Object.keys(groups).forEach(function(k){var a=groups[k],pos=a.filter(function(q){return q.amount>0;});if(!pos.length)return;pos.sort(function(a,b){return i74Rich_(b)-i74Rich_(a);});var rep=pos[0],orig=0,cancel=0,cancelDate='';a.forEach(function(q){if(q.amount>orig)orig=q.amount;var c=Math.abs(q.cancelAmount||0);if(q.amount<0)c=Math.max(c,Math.abs(q.amount));if(q.cancel&&q.amount>0)c=Math.max(c,q.amount);if(c>cancel){cancel=c;cancelDate=q.cancelDate||q.date;}});cancel=Math.min(cancel,orig);out.push(i74CO_(rep,Math.max(orig-cancel,0),orig,cancel,orig>0&&orig-cancel===0,k,cancelDate));});
  return out;
}
function i74Single_(q){if(q.cancel&&q.amount<0)return null;var orig=Math.abs(q.amount||0),cancel=Math.abs(q.cancelAmount||0);if(q.cancel&&orig>0)cancel=Math.max(cancel,orig);return i74CO_(q,Math.max(orig-cancel,0),orig,cancel,orig>0&&orig-cancel===0,'ROW|'+q.row,q.cancelDate);}
function i74CO_(q,amt,orig,cancel,full,k,cancelDate){var id=q.nonCard?'NONCARD|'+i74C_(q.cardName||q.evidence):'CARD|'+i74C_(q.company)+'|'+i74C_(q.cardNumber||q.end4||q.cardName);var key=q.approvalNo?'APP|'+i74C_(q.company)+'|'+i74C_(q.approvalNo):'ROW|'+q.date+'|'+amt+'|'+i74C_(q.merchant)+'|'+i74C_(q.source);return {company:q.company,cardName:q.cardName,cardNumber:q.cardNumber,end4:q.end4,date:q.date,time:q.time,merchant:q.merchant,amount:amt,originalAmount:orig,cancelAmount:cancel,cancelDate:cancelDate||'',approvalNo:q.approvalNo,merchantOrderNo:q.merchantOrderNo,evidence:q.evidence,source:q.source,nonCard:q.nonCard,lotteEvidence:q.lotteEvidence,fullyCanceled:full,identity:id,key:key,canonKey:k};}
function i74Rich_(q){var n=0;if(i74T_(q.evidence)==='카드이용내역')n+=100;if(i74T_(q.end4))n+=20;if(i74T_(q.cardNumber))n+=10;if(i74T_(q.cardName))n+=5;if(!q.cancel)n++;return n;}
function i74UsedSet_(rows){var u={};rows.forEach(function(o){if(o.status!=='MATCHED'&&o.status!=='NON_CARD')return;var k=i74EvidenceKey_(o.company,o.approvalNo,o.approvalDate,o.approvalAmount,o.merchant,o.source);if(k)u[k]=true;});return u;}
function i74RecoverDate_(o,h,used){var c=h.filter(function(q){var d=i74Days_(o.orderDate,q.date);return q.lotteEvidence&&!q.fullyCanceled&&q.amount===o.purchase&&Math.abs(d)<=14&&!(d>=0&&d<=7);});if(c.length!==1)return {ok:false,reason:'OUTSIDE_EXACT_CANDIDATES_'+c.length,summary:i74Summ_(c)};var q=c[0],lag=i74Days_(o.orderDate,q.date);if(lag!==-1)return {ok:false,reason:'LAG_'+lag,summary:i74Summ_(c)};if(q.key&&used[q.key])return {ok:false,reason:'EVIDENCE_USED',summary:i74Summ_(c)};if(i74PayConflict_(o.payment,q.company))return {ok:false,reason:'PAYMENT_CONFLICT',summary:i74Summ_(c)};return {ok:true,candidate:q,summary:i74Summ_(c)};}
function i74RecoverSplit_(o,h,used){var w=h.filter(function(q){var d=i74Days_(o.orderDate,q.date);return q.lotteEvidence&&!q.fullyCanceled&&d>=0&&d<=7;});var safe=[];for(var i=0;i<w.length;i++)for(var j=i+1;j<w.length;j++){var a=w[i],b=w[j];if(a.amount+b.amount!==o.purchase)continue;if(a.identity!==b.identity)continue;if((a.key&&used[a.key])||(b.key&&used[b.key]))continue;safe.push([a,b]);}if(safe.length!==1)return {ok:false,reason:'SAFE_PAIR_COUNT_'+safe.length,summary:i74Summ_(w)};var p=safe[0];if(String(p[0].identity).indexOf('NONCARD|')===0)return {ok:false,reason:'NONCARD_PAIR',summary:i74Summ_(p)};return {ok:true,pair:p,summary:i74Summ_(p)};}

function i74ApplySingle_(row,ix,o,q){i74Set_(row,ix.company,q.company);i74Set_(row,ix.alias,'');i74Set_(row,ix.cardName,q.cardName);i74Set_(row,ix.number,q.cardNumber);i74Set_(row,ix.end4,q.end4||i74End4_(q.cardNumber));i74Set_(row,ix.adate,q.date);i74Set_(row,ix.atime,q.time);i74Set_(row,ix.approval,q.approvalNo);i74Set_(row,ix.aamount,q.originalAmount||q.amount);i74Set_(row,ix.status,'MATCHED');i74Set_(row,ix.reason,'v6.74_-1일_exact_단일미사용증빙_1:1회수');i74Set_(row,ix.candidate,1);i74Set_(row,ix.merchant,q.merchant);i74Set_(row,ix.merchantOrder,q.merchantOrderNo);i74Set_(row,ix.evidence,q.evidence||'카드이용내역');i74Set_(row,ix.cancelMemo,q.cancelAmount?'취소금액 '+q.cancelAmount+' / effective '+q.amount:'');i74Set_(row,ix.source,q.source);i74Set_(row,ix.summary,i74Summ_([q]));}
function i74ApplySplit_(row,ix,o,p){var a=p[0],b=p[1],company=a.company,cardName=a.cardName||b.cardName,number=a.cardNumber||b.cardNumber,end4=a.end4||b.end4||i74End4_(number);i74Set_(row,ix.company,company);i74Set_(row,ix.alias,'');i74Set_(row,ix.cardName,cardName);i74Set_(row,ix.number,number);i74Set_(row,ix.end4,end4);i74Set_(row,ix.adate,a.date===b.date?a.date:a.date+'+'+b.date);i74Set_(row,ix.atime,'');i74Set_(row,ix.approval,'SPLIT:'+i74T_(a.approvalNo)+'+'+i74T_(b.approvalNo));i74Set_(row,ix.aamount,o.purchase);i74Set_(row,ix.status,'MATCHED');i74Set_(row,ix.reason,'v6.74_split_exact_동일카드_2증빙_1:1회수');i74Set_(row,ix.candidate,2);i74Set_(row,ix.merchant,a.merchant===b.merchant?a.merchant:a.merchant+' + '+b.merchant);i74Set_(row,ix.merchantOrder,'');i74Set_(row,ix.evidence,'분할결제_2증빙');i74Set_(row,ix.cancelMemo,'');i74Set_(row,ix.source,a.source===b.source?a.source:a.source+' + '+b.source);i74Set_(row,ix.summary,i74Summ_(p));}
function i74Set_(row,idx,val){if(idx>=0)row[idx]=val;}
function i74StatsFromMatrix_(v,hr,ix){var s={total:0,purchase:0,MATCHED:0,NON_CARD:0,AMBIGUOUS:0,NO_MATCH:0};for(var r=hr+1;r<v.length;r++){var no=i74T_(v[r][ix.order]);if(!no)continue;var st=i74T_(v[r][ix.status]).toUpperCase();s.total++;s.purchase+=i74N_(v[r][ix.purchase]);s[st]=(s[st]||0)+1;}s.purchase=Math.round(s.purchase);return s;}
function i74ChangedRows_(oldV,newV,hr,ix,eligible,recovered){var changed=0,unexpected=0,protectedChanged=0;for(var r=hr+1;r<oldV.length;r++){var no=i74T_(oldV[r][ix.order]);if(!no)continue;var same=JSON.stringify(oldV[r])===JSON.stringify(newV[r]);if(same)continue;changed++;var key=i74C_(oldV[r][ix.account])+'|'+i74C_(no),oldStatus=i74T_(oldV[r][ix.status]).toUpperCase();if(!eligible[key]||!recovered[key])unexpected++;if(oldStatus==='MATCHED'||oldStatus==='NON_CARD')protectedChanged++;}return {changed:changed,unexpected:unexpected,protectedChanged:protectedChanged};}
function i74WritePreview_(ss,values){var sh=ss.getSheetByName(I74_PREVIEW)||ss.insertSheet(I74_PREVIEW);if(sh.getFilter())sh.getFilter().remove();sh.clear();if(values.length&&values[0].length)sh.getRange(1,1,values.length,values[0].length).setValues(values);sh.setFrozenRows(1);try{sh.getRange(1,1,values.length,values[0].length).createFilter();}catch(_e){}}
function i74WriteRecoveryDetail_(ss,rows){var name='ISSUE74_회수변경내역',sh=ss.getSheetByName(name)||ss.insertSheet(name);if(sh.getFilter())sh.getFilter().remove();sh.clear();var h=['사업자등록번호','쿠팡계정ID','주문일','주문번호','주문매입금액','Issue73판정','Issue74결과','증빙일','카드사','카드identity','후보요약'];sh.getRange(1,1,1,h.length).setValues([h]);if(rows.length)sh.getRange(2,1,rows.length,h.length).setValues(rows);sh.setFrozenRows(1);try{sh.getRange(1,1,rows.length+1,h.length).createFilter();}catch(_e){}}
function i74Status_(ss,pairs){var sh=ss.getSheetByName(I74_STATUS)||ss.insertSheet(I74_STATUS);sh.clearContents();sh.getRange(1,1,1,2).setValues([['항목','값']]);if(pairs&&pairs.length)sh.getRange(2,1,pairs.length,2).setValues(pairs);sh.setFrozenRows(1);sh.getRange(1,1,1,2).setFontWeight('bold');sh.autoResizeColumns(1,2);}

function i74Header_(v,req){for(var r=0;r<Math.min(v.length,30);r++){var h=v[r]||[],ok=req.every(function(n){return i74Ix_(h,[n])>=0;});if(ok)return r;}return -1;}
function i74Ix_(h,names){var m={};(h||[]).forEach(function(x,i){m[i74C_(x)]=i;});for(var j=0;j<names.length;j++){var k=i74C_(names[j]);if(Object.prototype.hasOwnProperty.call(m,k))return m[k];}return -1;}
function i74FindHeaderContains_(h,needle){var n=i74C_(needle);for(var i=0;i<h.length;i++)if(i74C_(h[i]).indexOf(n)>=0)return i;return -1;}
function i74T_(v){if(v===null||v===undefined)return '';if(Object.prototype.toString.call(v)==='[object Date]')return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');return String(v).trim();}
function i74C_(v){return i74T_(v).toLowerCase().replace(/[\s_\-\/.()[\]:]+/g,'');}
function i74N_(v){if(typeof v==='number')return isNaN(v)?0:v;var s=String(v===null||v===undefined?'':v).replace(/,/g,'').replace(/[^0-9.\-]/g,'');var n=Number(s);return isNaN(n)?0:n;}
function i74D_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');var s=i74T_(v),m=s.match(/(20\d{2})[^0-9]?(\d{1,2})[^0-9]?(\d{1,2})/);return m?m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2):'';}
function i74Days_(a,b){var x=String(a||'').split('-'),y=String(b||'').split('-');if(x.length!==3||y.length!==3)return 99999;return Math.round((Date.UTC(+y[0],+y[1]-1,+y[2])-Date.UTC(+x[0],+x[1]-1,+x[2]))/86400000);}
function i74Cancel_(q){var s=i74C_(q.status);if(!s)return q.amount<0;if(s.indexOf('취소있음')>=0)return false;return s.indexOf('취소')>=0||s.indexOf('환불')>=0||q.amount<0;}
function i74PayConflict_(p,c){var a=i74C_(p),b=i74C_(c);if(!a||!b||/lpay|엘페이|토스|카카오/.test(a))return false;var k=[['국민','국민'],['kb','kb'],['우리','우리'],['신한','신한'],['농협','농협'],['nh','nh'],['삼성','삼성'],['하나','하나'],['현대','현대'],['롯데','롯데']];for(var i=0;i<k.length;i++)if(a.indexOf(k[i][0])>=0)return b.indexOf(k[i][1])<0;return false;}
function i74EvidenceKey_(c,a,d,m,mer,s){if(i74T_(a))return 'APP|'+i74C_(c)+'|'+i74C_(a);if(!d&&!m&&!mer&&!s)return '';return 'ROW|'+d+'|'+m+'|'+i74C_(mer)+'|'+i74C_(s);}
function i74End4_(n){return String(n||'').replace(/\D/g,'').slice(-4);}
function i74Summ_(a){return (a||[]).slice(0,8).map(function(q){return [q.date,q.amount,q.identity,q.approvalNo,q.key].join('|');}).join(' / ');}
function i74Sig_(sh){var r=sh.getLastRow(),c=sh.getLastColumn();if(r<1||c<1)return r+'x'+c+':EMPTY';var v=sh.getRange(1,1,r,c).getValues();var t=JSON.stringify(v,function(_k,x){if(Object.prototype.toString.call(x)==='[object Date]')return {__date__:x.getTime()};if(typeof x==='number'&&isNaN(x))return {__nan__:true};return x;});var b=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,t,Utilities.Charset.UTF_8);return r+'x'+c+':'+b.map(function(x){var n=(x<0?x+256:x).toString(16);return n.length===1?'0'+n:n;}).join('');}
