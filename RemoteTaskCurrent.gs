/**
 * Issue #56 v1.0 read-only adjudication for the 12 state-changed orders from Issue55.
 * Writes only ISSUE56_* sheets. Production VAT/card sheets remain read-only.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE56-v1.0-20260813',
  title: 'Issue55 상태변경 12건 카드증빙 자동 판정',
  enabled: true,
  outputSheet: 'ISSUE56_상태변경12건판정',
  statusSheet: 'ISSUE56_판정상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var status = i56Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  i56Write_(status, [
    ['항목','값'],['버전','v1.0-ISSUE56-CHANGED12-EVIDENCE-ADJUDICATION'],['상태','RUNNING'],['단계','LOAD'],
    ['메시지','Issue55 상태변경 12건 증빙 자동 판정 시작'],['운영시트 변경','0']
  ]);

  try {
    var deltaSh = ss.getSheetByName('ISSUE55_카드매칭차이진단');
    var previewSh = ss.getSheetByName('ISSUE54_카드매칭전체PREVIEW');
    var oldSh = ss.getSheetByName('부가세_카드매칭검증');
    var vatSh = ss.getSheetByName('부가세_신고자료');
    if (!deltaSh || deltaSh.getLastRow() < 2) throw new Error('ISSUE55_카드매칭차이진단이 없습니다.');
    if (!previewSh || previewSh.getLastRow() < 2) throw new Error('ISSUE54_카드매칭전체PREVIEW가 없습니다.');
    if (!oldSh || oldSh.getLastRow() < 2) throw new Error('부가세_카드매칭검증이 없습니다.');
    if (!vatSh || vatSh.getLastRow() < 2) throw new Error('부가세_신고자료가 없습니다.');

    var sigOld = i56Sig_(oldSh.getDataRange().getValues());
    var sigVat = i56Sig_(vatSh.getDataRange().getValues());
    var sigPreview = i56Sig_(previewSh.getDataRange().getValues());
    var sigDelta = i56Sig_(deltaSh.getDataRange().getValues());

    var pv = previewSh.getDataRange().getValues();
    var ph = pv[0].map(i56Text_);
    var px = {
      account:i56Find_(ph,['쿠팡계정ID']), order:i56Find_(ph,['주문번호']), purchase:i56Find_(ph,['주문매입금액','매입금액']),
      status:i56Find_(ph,['카드매칭상태']), reason:i56Find_(ph,['카드매칭근거']), company:i56Find_(ph,['구매카드사']),
      cardName:i56Find_(ph,['구매카드명']), end4:i56Find_(ph,['카드번호끝4']), approvalNo:i56Find_(ph,['승인번호']),
      approvalAmount:i56Find_(ph,['승인금액']), v669:i56Find_(ph,['v6.69 2차귀속']), v670:i56Find_(ph,['v6.70 3차귀속']),
      evidenceKey:i56Find_(ph,['canonicalEvidenceKey'])
    };
    i56Req_(px.account>=0 && px.order>=0 && px.purchase>=0 && px.status>=0 && px.reason>=0 && px.evidenceKey>=0,
      'Issue54 preview 필수 헤더 누락');

    var pCounts={MATCHED:0,NON_CARD:0,AMBIGUOUS:0,NO_MATCH:0};
    var pPurchase=0, keyOwners={}, duplicateMatchedKeys=0;
    for (var r=1;r<pv.length;r++) {
      var pr=pv[r], st=i56Status_(pr[px.status]);
      pCounts[st]=(pCounts[st]||0)+1;
      pPurchase += i56Num_(pr[px.purchase]);
      if (st==='MATCHED') {
        var ek=i56Text_(pr[px.evidenceKey]);
        if (ek) {
          var owner=i56Text_(pr[px.account]).toLowerCase()+'|'+i56NormOrder_(pr[px.order]);
          if (!keyOwners[ek]) keyOwners[ek]=[];
          keyOwners[ek].push(owner);
          if (keyOwners[ek].length===2) duplicateMatchedKeys++;
        }
      }
    }
    i56Req_(pv.length-1===1355,'Issue54 preview 주문수 불일치: '+(pv.length-1));
    i56Req_(pCounts.MATCHED===808 && pCounts.NON_CARD===498 && pCounts.AMBIGUOUS===0 && pCounts.NO_MATCH===49,
      'Issue54 상태 기준 불일치: '+JSON.stringify(pCounts));
    i56Req_(Math.round(pPurchase)===105762969,'Issue54 매입합계 불일치: '+Math.round(pPurchase));
    i56Req_(duplicateMatchedKeys===0,'Issue54 MATCHED canonical key 중복 사용: '+duplicateMatchedKeys);

    var dv=deltaSh.getDataRange().getValues(), dh=dv[0].map(i56Text_);
    var dx={
      account:i56Find_(dh,['쿠팡계정ID']), order:i56Find_(dh,['주문번호']), date:i56Find_(dh,['주문일']),
      oldStatus:i56Find_(dh,['기존상태']), newStatus:i56Find_(dh,['신규상태']), transition:i56Find_(dh,['상태이동']),
      oldPurchase:i56Find_(dh,['기존매입금액']), newPurchase:i56Find_(dh,['신규매입금액']), diff:i56Find_(dh,['매입차액']),
      oldCompany:i56Find_(dh,['기존카드사']), newCompany:i56Find_(dh,['신규카드사']), oldCard:i56Find_(dh,['기존카드명']), newCard:i56Find_(dh,['신규카드명']),
      oldEnd4:i56Find_(dh,['기존끝4']), newEnd4:i56Find_(dh,['신규끝4']), oldApprovalNo:i56Find_(dh,['기존승인번호']), newApprovalNo:i56Find_(dh,['신규승인번호']),
      oldApprovalAmount:i56Find_(dh,['기존승인금액']), newApprovalAmount:i56Find_(dh,['신규승인금액']), oldReason:i56Find_(dh,['기존매칭근거']), newReason:i56Find_(dh,['신규매칭근거']),
      old669:i56Find_(dh,['기존v6.69']), new669:i56Find_(dh,['신규v6.69']), old670:i56Find_(dh,['기존v6.70']), new670:i56Find_(dh,['신규v6.70']),
      oldKey:i56Find_(dh,['기존canonicalKey']), newKey:i56Find_(dh,['신규canonicalKey'])
    };
    i56Req_(dx.account>=0&&dx.order>=0&&dx.oldStatus>=0&&dx.newStatus>=0&&dx.transition>=0&&dx.oldKey>=0&&dx.newKey>=0,
      'Issue55 delta 필수 헤더 누락');
    i56Req_(dv.length-1===12,'Issue55 상태변경 주문수 불일치: '+(dv.length-1));

    var expectedTransitions={'AMBIGUOUS -> MATCHED':1,'MATCHED -> NO_MATCH':5,'NO_MATCH -> MATCHED':2,'NO_MATCH -> NON_CARD':4};
    var trans={}, outRows=[], autoSafe=0, review=0, invalid=0, purchaseChanged=0;
    var safeMatched=0, safeNonCard=0, lostReassigned=0, lostUnused=0, lostFallback=0;

    for (var i=1;i<dv.length;i++) {
      var row=dv[i];
      var account=i56Text_(row[dx.account]).toLowerCase();
      var order=i56Text_(row[dx.order]);
      var orderKey=account+'|'+i56NormOrder_(order);
      var tr=i56Text_(row[dx.transition]);
      var os=i56Status_(row[dx.oldStatus]), ns=i56Status_(row[dx.newStatus]);
      var oldKey=i56Text_(row[dx.oldKey]), newKey=i56Text_(row[dx.newKey]);
      var oldReason=i56Text_(row[dx.oldReason]), newReason=i56Text_(row[dx.newReason]);
      var old669=i56Yes_(row[dx.old669]), new669=i56Yes_(row[dx.new669]), old670=i56Yes_(row[dx.old670]), new670=i56Yes_(row[dx.new670]);
      var newCompany=i56Text_(row[dx.newCompany]), newCard=i56Text_(row[dx.newCard]), newEnd4=i56Text_(row[dx.newEnd4]);
      var newApprovalNo=i56Text_(row[dx.newApprovalNo]), newApprovalAmount=i56Num_(row[dx.newApprovalAmount]);
      var oldP=i56Num_(row[dx.oldPurchase]), newP=i56Num_(row[dx.newPurchase]);
      if (Math.round(oldP)!==Math.round(newP)) purchaseChanged++;
      trans[tr]=(trans[tr]||0)+1;

      var verdict='', detail='', related='';
      if (ns==='MATCHED') {
        var keyed = !!newKey;
        var fallback = new669 || new670;
        var keyUse = keyed && keyOwners[newKey] ? keyOwners[newKey].length : 0;
        if (!newReason) {
          verdict='INVALID'; detail='신규 MATCHED인데 매칭근거 공란'; invalid++;
        } else if (keyed && keyUse===1) {
          verdict='AUTO_SAFE'; detail='신규 MATCHED canonical key 1:1 사용'; autoSafe++; safeMatched++;
        } else if (!keyed && fallback && (newCompany||newCard||newEnd4) && newApprovalAmount>=0) {
          verdict='AUTO_SAFE'; detail='신규 MATCHED fallback 근거 및 카드 identity 존재'; autoSafe++; safeMatched++;
        } else {
          verdict='REVIEW_REQUIRED'; detail='신규 MATCHED 증빙 형태 추가 확인 필요 (keyUse='+keyUse+', fallback='+(fallback?'Y':'N')+')'; review++;
        }
      } else if (os==='MATCHED' && ns==='NO_MATCH') {
        if (newKey || newCompany || newCard || newEnd4 || newApprovalNo) {
          verdict='INVALID'; detail='신규 NO_MATCH인데 카드 identity/canonical key 잔존'; invalid++;
        } else if (oldKey && keyOwners[oldKey] && keyOwners[oldKey].length===1 && keyOwners[oldKey][0]!==orderKey) {
          verdict='AUTO_SAFE'; related=keyOwners[oldKey][0]; detail='기존 증빙이 신규 다른 주문에 1:1 재배정됨'; autoSafe++; lostReassigned++;
        } else if (!oldKey && (old669||old670)) {
          verdict='REVIEW_REQUIRED'; detail='기존 fallback MATCHED가 신규 NO_MATCH로 해제됨'; review++; lostFallback++;
        } else {
          verdict='REVIEW_REQUIRED'; detail='기존 MATCHED 증빙이 신규 preview에서 재사용되지 않음'; review++; lostUnused++;
        }
      } else if (ns==='NON_CARD') {
        if (!newReason) {
          verdict='INVALID'; detail='신규 NON_CARD인데 근거 공란'; invalid++;
        } else if (newKey || newEnd4 || newApprovalNo) {
          verdict='INVALID'; detail='신규 NON_CARD인데 physical card 증빙 잔존'; invalid++;
        } else {
          verdict='AUTO_SAFE'; detail='신규 NON_CARD 근거 존재 및 physical card 증빙 없음'; autoSafe++; safeNonCard++;
        }
      } else {
        verdict='REVIEW_REQUIRED'; detail='예상하지 않은 상태이동'; review++;
      }

      outRows.push([
        account,order,i56Text_(row[dx.date]),tr,oldP,newP,Math.round(newP-oldP),verdict,detail,related,
        oldKey,newKey,oldReason,newReason,old669?'Y':'',new669?'Y':'',old670?'Y':'',new670?'Y':'',
        i56Text_(row[dx.oldCompany]),newCompany,i56Text_(row[dx.oldCard]),newCard,i56Text_(row[dx.oldEnd4]),newEnd4,
        i56Text_(row[dx.oldApprovalNo]),newApprovalNo,i56Num_(row[dx.oldApprovalAmount]),newApprovalAmount
      ]);
    }

    Object.keys(expectedTransitions).forEach(function(k){i56Req_((trans[k]||0)===expectedTransitions[k],'상태이동 기준 불일치 '+k+': '+(trans[k]||0));});
    i56Req_(Object.keys(trans).length===4,'예상 외 상태이동 종류 존재: '+JSON.stringify(trans));
    i56Req_(purchaseChanged===8,'상태변경 중 매입금액변경 기준 불일치: '+purchaseChanged);

    var out=i56Ensure_(ss,LOTTEON_REMOTE_TASK.outputSheet);
    out.clearContents();
    var headers=['쿠팡계정ID','주문번호','주문일','상태이동','기존매입금액','신규매입금액','매입차액','자동판정','판정근거','관련신규주문','기존canonicalKey','신규canonicalKey','기존매칭근거','신규매칭근거','기존v6.69','신규v6.69','기존v6.70','신규v6.70','기존카드사','신규카드사','기존카드명','신규카드명','기존끝4','신규끝4','기존승인번호','신규승인번호','기존승인금액','신규승인금액'];
    out.getRange(1,1,1,headers.length).setValues([headers]);
    if (outRows.length) out.getRange(2,1,outRows.length,headers.length).setValues(outRows);
    out.getRange(1,1,1,headers.length).setFontWeight('bold');
    out.setFrozenRows(1);
    SpreadsheetApp.flush();

    i56Req_(i56Sig_(oldSh.getDataRange().getValues())===sigOld,'부가세_카드매칭검증이 판정 중 변경되었습니다.');
    i56Req_(i56Sig_(vatSh.getDataRange().getValues())===sigVat,'부가세_신고자료가 판정 중 변경되었습니다.');
    i56Req_(i56Sig_(previewSh.getDataRange().getValues())===sigPreview,'Issue54 preview가 판정 중 변경되었습니다.');
    i56Req_(i56Sig_(deltaSh.getDataRange().getValues())===sigDelta,'Issue55 delta가 판정 중 변경되었습니다.');

    var rows=[
      ['항목','값'],['버전','v1.0-ISSUE56-CHANGED12-EVIDENCE-ADJUDICATION'],['상태','PASS'],['단계','DONE'],
      ['메시지','Issue55 상태변경 12건 증빙 자동 판정 완료'],['운영시트 변경','0'],
      ['상태변경주문',12],['AUTO_SAFE',autoSafe],['REVIEW_REQUIRED',review],['INVALID',invalid],
      ['신규MATCHED_AUTO_SAFE',safeMatched],['신규NON_CARD_AUTO_SAFE',safeNonCard],
      ['기존MATCHED증빙_다른주문재배정',lostReassigned],['기존MATCHED증빙_재사용없음',lostUnused],['기존fallbackMATCHED_해제',lostFallback],
      ['상태변경중_매입금액변경',purchaseChanged],['Issue54_MATCHED_key중복',duplicateMatchedKeys],
      ['이동_AMBIGUOUS -> MATCHED',trans['AMBIGUOUS -> MATCHED']||0],['이동_MATCHED -> NO_MATCH',trans['MATCHED -> NO_MATCH']||0],
      ['이동_NO_MATCH -> MATCHED',trans['NO_MATCH -> MATCHED']||0],['이동_NO_MATCH -> NON_CARD',trans['NO_MATCH -> NON_CARD']||0],
      ['운영반영자동승인',review===0&&invalid===0?'YES':'NO'],
      ['부가세_카드매칭검증 변경','0'],['부가세_신고자료 변경','0'],['완료시각',new Date().toISOString()]
    ];
    for (var q=0;q<outRows.length;q++) {
      if (outRows[q][7]!=='AUTO_SAFE') rows.push(['확인필요_'+(q+1),outRows[q][0]+' | '+outRows[q][1]+' | '+outRows[q][3]+' | '+outRows[q][7]+' | '+outRows[q][8]]);
    }
    i56Write_(status,rows);
    return {ok:true,done:true,autoSafe:autoSafe,reviewRequired:review,invalid:invalid};
  } catch(e) {
    i56Write_(status,[
      ['항목','값'],['버전','v1.0-ISSUE56-CHANGED12-EVIDENCE-ADJUDICATION'],['상태','ERROR'],['단계','FAILED'],
      ['메시지','Issue55 상태변경 12건 증빙 자동 판정 실패'],['오류',String(e&&e.message?e.message:e)],['운영시트 변경','0'],['완료시각',new Date().toISOString()]
    ]);
    throw e;
  }
}

function runLotteonRemoteTaskContinueRemote_(){return {ok:true,done:true,reason:'NO_CONTINUE_REQUIRED'};}
function i56Status_(v){var s=i56Text_(v).toUpperCase();if(s==='MATCHED'||s==='MASTER_MATCHED')return 'MATCHED';if(s==='NON_CARD')return 'NON_CARD';if(s==='AMBIGUOUS')return 'AMBIGUOUS';return 'NO_MATCH';}
function i56Yes_(v){var s=i56Text_(v).toUpperCase();return s==='Y'||s==='YES'||s==='TRUE'||s==='1';}
function i56NormOrder_(v){return i56Text_(v).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');}
function i56Find_(h,names){for(var n=0;n<names.length;n++){var q=i56Compact_(names[n]);for(var i=0;i<h.length;i++)if(i56Compact_(h[i])===q)return i;}return -1;}
function i56Text_(v){return String(v==null?'':v).trim();}
function i56Compact_(v){return i56Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function i56Num_(v){if(typeof v==='number'&&isFinite(v))return v;var n=Number(i56Text_(v).replace(/[원,%\s,]/g,''));return isFinite(n)?n:0;}
function i56Req_(ok,msg){if(!ok)throw new Error(msg);}
function i56Ensure_(ss,name){return ss.getSheetByName(name)||ss.insertSheet(name);}
function i56Write_(sh,rows){sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setFrozenRows(1);SpreadsheetApp.flush();}
function i56Sig_(v){var h=2166136261;for(var r=0;r<v.length;r++){for(var c=0;c<v[r].length;c++){var s=String(v[r][c]==null?'':v[r][c]);for(var i=0;i<s.length;i++){h^=s.charCodeAt(i);h+=(h<<1)+(h<<4)+(h<<7)+(h<<8)+(h<<24);}h^=31;}}return String(h>>>0)+'|'+v.length+'|'+(v[0]?v[0].length:0);}
