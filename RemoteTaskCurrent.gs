var LOTTEON_REMOTE_TASK={id:'ISSUE79-NOMATCH16-DEEP-v4',title:'v2 SAFE 6 보존 + 잔여16 카드힌트 심층검수',enabled:true,statusSheet:'ISSUE79_심층재검수상태'};
var ISSUE79_V2_SOURCE='https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/d63df776c5f0fb7232e2fdac17558fa58668a1fa/RemoteTaskCurrent.gs';
function runLotteonRemoteTaskStartRemote_(){return issue79V4Exec_();}
function runLotteonRemoteTaskContinueRemote_(){return issue79V4Exec_();}
function issue79V4Exec_(){
  var res=UrlFetchApp.fetch(ISSUE79_V2_SOURCE+'?ts='+new Date().getTime(),{method:'get',muteHttpExceptions:true,followRedirects:true});
  var code=res.getContentText();
  if(res.getResponseCode()!==200||!code)throw new Error('Issue79 v2 source fetch 실패 HTTP '+res.getResponseCode());
  var extra=String.raw`
function i79v4deep_(){
  var ss=SpreadsheetApp.getActive(),core=['매출데이터_붙여넣기','부가세_신고자료','부가세_카드매칭검증','부가세_기간별','카드사용내역_붙여넣기','카드_마스터'],before={};
  core.forEach(function(n){var sh=ss.getSheetByName(n);if(!sh)throw new Error('핵심 시트 누락 '+n);before[n]=i79sig_(sh);});
  var pv=ss.getSheetByName('ISSUE79_NOMATCH22_최신재검수');if(!pv)throw new Error('v2 preview 시트 누락');
  var vr=pv.getDataRange(),vv=vr.getValues(),vd=vr.getDisplayValues(),h=vv[0];
  var x={order:i79ix_(h,['주문번호']),targetType:i79ix_(h,['재매칭기준']),targetAmount:i79ix_(h,['재매칭목표금액']),verdict:i79ix_(h,['PREVIEW판정']),bestCompany:i79ix_(h,['후보카드사']),bestEnd4:i79ix_(h,['후보끝4']),bestApproval:i79ix_(h,['후보승인번호']),bestDate:i79ix_(h,['후보승인일']),bestAmount:i79ix_(h,['후보금액'])};
  Object.keys(x).forEach(function(k){if(x[k]<0)throw new Error('v2 preview header 누락 '+k);});
  var v2={},safeOrders=[],remainOrders=[];for(var r=1;r<vv.length;r++){var no=i79t_(vd[r][x.order]);if(!no)continue;var z={row:r+1,orderNo:no,targetType:i79t_(vv[r][x.targetType]),targetAmount:i79n_(vv[r][x.targetAmount]),verdict:i79t_(vv[r][x.verdict]),company:i79t_(vv[r][x.bestCompany]),end4:i79t_(vd[r][x.bestEnd4]),approvalNo:i79t_(vd[r][x.bestApproval]),date:i79date_(vv[r][x.bestDate]),amount:i79n_(vv[r][x.bestAmount])};v2[no]=z;if(/^SAFE_/.test(z.verdict))safeOrders.push(no);else remainOrders.push(no);}
  if(safeOrders.length!==6||remainOrders.length!==16)throw new Error('v2 SAFE/잔여 guard 실패 safe='+safeOrders.length+' remain='+remainOrders.length);
  var card=i79readCard_(ss.getSheetByName('부가세_카드매칭검증'));if((card.stats.NO_MATCH||0)!==22)throw new Error('운영 NO_MATCH guard '+JSON.stringify(card.stats));
  var byOrder={};card.rows.forEach(function(o){byOrder[o.orderNo]=o;});
  var srcHints=i79v4sourceHints_(ss.getSheetByName('매출데이터_붙여넣기'));
  var hist=i79canon_(i79readHist_(ss.getSheetByName('카드사용내역_붙여넣기'))),used=i79used_(card.rows),reserved={};
  safeOrders.forEach(function(no){var z=v2[no],k=z.approvalNo?'APP|'+i79c_(z.company)+'|'+i79c_(z.approvalNo):'';if(k)reserved[k]=1;});
  var out=[],sum={v2Safe:6,deepSafe:0,noHint:0,hintRows:0,approvalExact:0,dateCardExact:0,cardWideExact:0,multi:0,used:0,noCandidate:0,blockedIntlMissing:0,hintConflict:0,cardHint:0,sourceHint:0};
  remainOrders.forEach(function(no){
    var base=v2[no],o=byOrder[no];if(!o)throw new Error('운영 주문 누락 '+no);var sh=srcHints.map[no]||{};
    var merged=i79v4mergeHints_(o,sh);if(merged.conflict)sum.hintConflict++;
    var hasStrong=!!(merged.approvalNo||merged.end4||merged.company||merged.cardName||merged.approvalDate);if(hasStrong)sum.hintRows++;else sum.noHint++;
    if(merged.fromCard)sum.cardHint++;if(merged.fromSource)sum.sourceHint++;
    var verdict='',reason='',cand=[];
    if(base.verdict==='CANCEL_INTL_FEE_MISSING'){verdict='BLOCKED_INTL_FEE_MISSING';sum.blockedIntlMissing++;}
    else if(merged.conflict){verdict='HINT_CONFLICT';reason=merged.conflictText;}
    else if(!hasStrong){verdict='NO_HINT';}
    else{
      cand=hist.filter(function(q){
        if(q.nonCard||q.fullyCanceled||!q.lotteEvidence||q.amount!==base.targetAmount)return false;
        var key=q.key||'';if(key&&(used[key]||reserved[key]))return false;
        if(merged.approvalNo){if(i79c_(q.approvalNo)!==i79c_(merged.approvalNo))return false;return i79v4cardCompatible_(merged,q);}
        if(merged.approvalDate&&q.date!==merged.approvalDate)return false;
        if(!i79v4cardCompatible_(merged,q))return false;
        var lag=i79days_(o.orderDate,q.date);return Math.abs(lag)<=90;
      });
      cand.sort(function(a,b){return Math.abs(i79days_(o.orderDate,a.date))-Math.abs(i79days_(o.orderDate,b.date));});
      if(cand.length===1){
        var q=cand[0];verdict='SAFE_DEEP_HINT_EXACT';sum.deepSafe++;reserved[q.key]=1;
        if(merged.approvalNo){reason='APPROVAL_NO_EXACT';sum.approvalExact++;}
        else if(merged.approvalDate){reason='APPROVAL_DATE_CARD_EXACT';sum.dateCardExact++;}
        else{reason='CARD_HINT_WIDE_EXACT';sum.cardWideExact++;}
      }else if(cand.length>1){verdict='MULTI_CANDIDATE';sum.multi++;}
      else{
        var usedCand=hist.filter(function(q){if(q.nonCard||q.fullyCanceled||!q.lotteEvidence||q.amount!==base.targetAmount)return false;if(!i79v4cardCompatible_(merged,q))return false;if(merged.approvalNo&&i79c_(q.approvalNo)!==i79c_(merged.approvalNo))return false;if(merged.approvalDate&&q.date!==merged.approvalDate)return false;var lag=i79days_(o.orderDate,q.date);return merged.approvalNo||merged.approvalDate||Math.abs(lag)<=90;}).filter(function(q){return q.key&&(used[q.key]||reserved[q.key]);});
        if(usedCand.length){verdict='USED_OR_RESERVED';sum.used++;cand=usedCand;}else{verdict='NO_CANDIDATE';sum.noCandidate++;}
      }
    }
    var best=cand[0]||{};
    out.push([no,o.orderDate,base.targetType,base.targetAmount,base.verdict,merged.company||'',merged.cardName||'',merged.end4||'',merged.approvalDate||'',merged.approvalNo||'',merged.approvalAmount||'',merged.fromCard?'Y':'',merged.fromSource?'Y':'',verdict,reason,cand.length,best.date||'',best.company||'',best.cardName||'',best.end4||'',best.approvalNo||'',best.amount||'',best.merchant||'',best.source||'',i79summary_(cand)]);
  });
  if(out.length!==16)throw new Error('심층진단 행수 '+out.length);
  var sh=ss.getSheetByName('ISSUE79_NOMATCH16_심층진단')||ss.insertSheet('ISSUE79_NOMATCH16_심층진단');if(sh.getFilter())sh.getFilter().remove();sh.clear();
  var hh=['주문번호','주문일','v2재매칭기준','목표금액','v2판정','힌트카드사','힌트카드명','힌트끝4','힌트승인일','힌트승인번호','힌트승인금액','운영카드힌트','매출원본힌트','v4판정','v4근거','후보수','후보승인일','후보카드사','후보카드명','후보끝4','후보승인번호','후보금액','후보가맹점','후보원본','후보요약'];sh.getRange(1,1,1,hh.length).setValues([hh]);if(out.length)sh.getRange(2,1,out.length,hh.length).setValues(out);sh.setFrozenRows(1);sh.getRange(1,1,1,hh.length).setFontWeight('bold');sh.getRange(2,1,out.length,1).setNumberFormat('@');sh.getRange(2,8,out.length,1).setNumberFormat('@');sh.getRange(2,10,out.length,1).setNumberFormat('@');sh.getRange(2,20,out.length,2).setNumberFormat('@');try{sh.getRange(1,1,out.length+1,hh.length).createFilter();}catch(_e){}
  var changed=[];core.forEach(function(n){var q=ss.getSheetByName(n);if(i79sig_(q)!==before[n])changed.push(n);});if(changed.length)throw new Error('READ-ONLY 위반 '+changed.join(','));
  var totalSafe=sum.v2Safe+sum.deepSafe,remain=22-totalSafe;
  var st=ss.getSheetByName('ISSUE79_심층재검수상태')||ss.insertSheet('ISSUE79_심층재검수상태');st.clearContents();var p=[['항목','값'],['version','v4.0-ISSUE79-V2SAFE-PRESERVE-DEEP-HINT-READONLY'],['상태','PASS'],['단계','DONE'],['메시지','v2 SAFE 6건 보존 + 잔여 16건 카드힌트 심층 READ-ONLY 재검수 완료'],['NO_MATCH기준',22],['v2_SAFE_보존',sum.v2Safe],['심층검수대상',16],['운영카드힌트존재행',sum.cardHint],['매출원본힌트존재행',sum.sourceHint],['강한힌트존재행',sum.hintRows],['힌트없음',sum.noHint],['힌트충돌',sum.hintConflict],['DEEP_SAFE_APPROVAL_NO_EXACT',sum.approvalExact],['DEEP_SAFE_APPROVAL_DATE_CARD_EXACT',sum.dateCardExact],['DEEP_SAFE_CARD_HINT_WIDE_EXACT',sum.cardWideExact],['DEEP_SAFE_추가',sum.deepSafe],['최종SAFE_합계',totalSafe],['MULTI_CANDIDATE',sum.multi],['USED_OR_RESERVED',sum.used],['NO_CANDIDATE',sum.noCandidate],['BLOCKED_INTL_FEE_MISSING',sum.blockedIntlMissing],['최종잔여검토',remain],['핵심시트변경수',0],['오류',''],['완료시각',new Date().toISOString()]];st.getRange(1,1,p.length,2).setValues(p);st.setFrozenRows(1);st.getRange(1,1,1,2).setFontWeight('bold');st.autoResizeColumns(1,2);
  return{ok:true,done:true,version:'v4.0-ISSUE79-V2SAFE-PRESERVE-DEEP-HINT-READONLY',v2Safe:sum.v2Safe,deepSafe:sum.deepSafe,totalSafe:totalSafe,remain:remain};
}
function i79v4sourceHints_(sh){
  var v=sh.getDataRange().getValues(),d=sh.getDataRange().getDisplayValues(),hr=-1,x=null,headers={};
  for(var r=0;r<Math.min(v.length,30);r++){var h=v[r]||[],oi=i79ix_(h,['주문번호','마켓주문번호','상품주문번호']);if(oi<0)continue;var z={order:oi,company:i79ix_(h,['구매카드사','카드사','결제카드사']),cardName:i79ix_(h,['구매카드명','카드명','결제카드명']),number:i79ix_(h,['구매카드번호','카드번호','결제카드번호']),end4:i79ix_(h,['카드번호끝4','카드끝4','끝4']),approvalDate:i79ix_(h,['승인일','카드승인일','결제승인일']),approvalNo:i79ix_(h,['승인번호','카드승인번호','결제승인번호']),approvalAmount:i79ix_(h,['승인금액','카드승인금액','결제승인금액'])};if(z.company>=0||z.cardName>=0||z.number>=0||z.end4>=0||z.approvalDate>=0||z.approvalNo>=0){hr=r;x=z;Object.keys(z).forEach(function(k){if(z[k]>=0)headers[k]=i79t_(h[z[k]]);});break;}}
  var m={};if(hr<0)return{map:m,headers:headers};for(var i=hr+1;i<v.length;i++){var no=i79t_(d[i][x.order]);if(!no)continue;var q=m[no]||{};function put(k,val){if(val&&!q[k])q[k]=val;}put('company',x.company>=0?i79t_(v[i][x.company]):'');put('cardName',x.cardName>=0?i79t_(v[i][x.cardName]):'');var num=x.number>=0?i79t_(d[i][x.number]):'';put('number',num);put('end4',x.end4>=0?i79t_(d[i][x.end4]):(num?String(num).replace(/\D/g,'').slice(-4):''));put('approvalDate',x.approvalDate>=0?i79date_(v[i][x.approvalDate]):'');put('approvalNo',x.approvalNo>=0?i79t_(d[i][x.approvalNo]):'');put('approvalAmount',x.approvalAmount>=0?i79n_(v[i][x.approvalAmount]):0);m[no]=q;}return{map:m,headers:headers};
}
function i79v4mergeHints_(o,s){
  var q={company:o.company||'',cardName:o.cardName||'',end4:o.end4||'',approvalDate:o.approvalDate||'',approvalNo:o.approvalNo||'',approvalAmount:o.approvalAmount||0,fromCard:!!(o.company||o.cardName||o.end4||o.approvalDate||o.approvalNo),fromSource:false,conflict:false,conflictText:''};
  s=s||{};var conflicts=[];function merge(k,v,cmp){if(!v)return;if(q[k]){if(cmp&&!cmp(q[k],v))conflicts.push(k+':'+q[k]+'<>'+v);}else{q[k]=v;q.fromSource=true;}}
  merge('company',s.company,i79v4companyCompat_);merge('cardName',s.cardName,function(a,b){return i79c_(a)===i79c_(b)||i79c_(a).indexOf(i79c_(b))>=0||i79c_(b).indexOf(i79c_(a))>=0;});merge('end4',s.end4,function(a,b){return i79v4last4_(a)===i79v4last4_(b);});merge('approvalDate',s.approvalDate,function(a,b){return i79date_(a)===i79date_(b);});merge('approvalNo',s.approvalNo,function(a,b){return i79c_(a)===i79c_(b);});if(!q.approvalAmount&&s.approvalAmount){q.approvalAmount=s.approvalAmount;q.fromSource=true;}q.conflict=conflicts.length>0;q.conflictText=conflicts.join(' / ');return q;
}
function i79v4last4_(v){var s=String(v==null?'':v).replace(/\D/g,'');return s.length>=4?s.slice(-4):s;}
function i79v4companyCompat_(a,b){var x=i79c_(a),y=i79c_(b);if(!x||!y)return true;var g=[['kb','국민'],['nh','농협'],['신한'],['우리'],['삼성'],['하나'],['현대'],['롯데']];for(var i=0;i<g.length;i++){var ax=g[i].some(function(t){return x.indexOf(t)>=0;});if(ax)return g[i].some(function(t){return y.indexOf(t)>=0;});}return x===y||x.indexOf(y)>=0||y.indexOf(x)>=0;}
function i79v4cardCompatible_(h,q){if(h.company&&!i79v4companyCompat_(h.company,q.company))return false;var e=i79v4last4_(h.end4);if(e&&i79v4last4_(q.end4)!==e)return false;if(!e&&!h.company&&h.cardName){var a=i79c_(h.cardName),b=i79c_(q.cardName);if(a&&b&&a!==b&&a.indexOf(b)<0&&b.indexOf(a)<0)return false;}return true;}
`;
  return eval(code+'\n'+extra+'\n;i79runV2_();\ni79v4deep_();');
}
