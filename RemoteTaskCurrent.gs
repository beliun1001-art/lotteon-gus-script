/**
 * Issue #51 v1.0 read-only full card rematch preview.
 * - Current 2026 Apr-Jun VAT: 1,893 orders / 3,894 detail rows.
 * - Reuses old verified LOTTE payment text only for exact normalized overlap keys.
 * - New/blank-payment orders stay blank; no payment inference.
 * - Uses current main Code.gs + v6.72 bootstrap (v6.60..v6.70 card chain).
 * - 150-order resumable batches; writes only ISSUE51_* sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE51-v1.0-20260812',
  title: '현재 VAT 1,893주문 카드매칭 전체 preview',
  enabled: true,
  outputSheet: 'ISSUE51_카드매칭전체PREVIEW',
  statusSheet: 'ISSUE51_실행상태'
};
var ISSUE51_V1_BATCH = 150;
var ISSUE51_V1_STATE_KEY = 'ISSUE51_V1_STATE';
var ISSUE51_V1_CODE_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Code.gs';
var ISSUE51_V1_BOOTSTRAP_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_24_bootstrap_auto_continue.gs';

function runLotteonRemoteTaskStartRemote_() {
  var props = PropertiesService.getScriptProperties();
  var prior = issue51v1ReadState_(props);
  if (prior && prior.taskId === LOTTEON_REMOTE_TASK.id) {
    if (prior.done) return {ok:true,done:true,reason:'ALREADY_DONE'};
    if (prior.terminalError) return {ok:false,done:true,reason:'TERMINAL_ERROR'};
    return runLotteonRemoteTaskContinueRemote_();
  }

  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  issue51v1DeleteContinueTriggers_();

  var status = issue51v1Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue51v1WriteStatus_(status, [
    ['항목','값'],['버전','v1.0-ISSUE51-CARD-REMATCH-PREVIEW'],['상태','RUNNING'],['단계','PRECHECK'],
    ['메시지','현재 VAT 1,893주문 전체 카드매칭 preview 사전검증 시작'],['운영시트 변경','0'],
    ['갱신시각',new Date().toISOString()]
  ]);

  try {
    var pre = issue51v1Precheck_(ss);
    issue51v1Assert_(pre.vatRows === 3894, '현재 VAT 상세행 불일치: ' + pre.vatRows);
    issue51v1Assert_(pre.vatOrders === 1893, '현재 VAT 주문수 불일치: ' + pre.vatOrders);
    issue51v1Assert_(Math.round(pre.vatPurchase) === 106707957, '현재 VAT 매입합계 불일치: ' + Math.round(pre.vatPurchase));
    issue51v1Assert_(pre.verifyOrders === 1355, '기존검증 주문수 불일치: ' + pre.verifyOrders);
    issue51v1Assert_(pre.overlap === 1355, '기존검증 겹침 주문수 불일치: ' + pre.overlap);
    issue51v1Assert_(pre.currentOnly === 538, '신규 주문수 불일치: ' + pre.currentOnly);
    issue51v1Assert_(pre.paymentReuse === 1270, '기존 결제수단 재사용 가능 주문수 불일치: ' + pre.paymentReuse);
    issue51v1Assert_(pre.verifyDup === 0 && pre.vatNormCollisions === 0, '정규화키 충돌/중복 존재');

    var out = issue51v1Ensure_(ss, LOTTEON_REMOTE_TASK.outputSheet);
    out.clearContents();
    var headers = issue51v1Headers_();
    out.getRange(1,1,1,headers.length).setValues([headers]);
    out.getRange(1,1,1,headers.length).setFontWeight('bold');
    out.setFrozenRows(1);

    var state = {
      taskId:LOTTEON_REMOTE_TASK.id,
      spreadsheetId:ss.getId(),
      startedAt:new Date().toISOString(),
      done:false,
      terminalError:false
    };
    props.setProperty(ISSUE51_V1_STATE_KEY, JSON.stringify(state));
    issue51v1WriteStatus_(status, [
      ['항목','값'],['버전','v1.0-ISSUE51-CARD-REMATCH-PREVIEW'],['상태','RUNNING'],['단계','BATCH'],
      ['메시지','사전검증 PASS; 1차 배치 시작'],['운영시트 변경','0'],
      ['처리주문',0],['대상주문',1893],['배치크기',ISSUE51_V1_BATCH],
      ['기존검증겹침',1355],['신규주문',538],['기존결제수단재사용가능',1270],
      ['시작시각',state.startedAt],['갱신시각',new Date().toISOString()]
    ]);
    return issue51v1ContinueCore_();
  } catch (e) {
    issue51v1Fail_(ss, e);
    throw e;
  }
}

function runLotteonRemoteTaskContinueRemote_() {
  return issue51v1ContinueCore_();
}

function issue51v1ContinueCore_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return {ok:false,reason:'LOCK_BUSY'};
  try {
    var props = PropertiesService.getScriptProperties();
    var st = issue51v1ReadState_(props);
    if (!st || st.taskId !== LOTTEON_REMOTE_TASK.id) throw new Error('ISSUE51 실행 상태가 없습니다. 다시 시작하세요.');
    if (st.done) return {ok:true,done:true,reason:'ALREADY_DONE'};
    if (st.terminalError) return {ok:false,done:true,reason:'TERMINAL_ERROR'};

    var ss = SpreadsheetApp.openById(st.spreadsheetId);
    var out = ss.getSheetByName(LOTTEON_REMOTE_TASK.outputSheet);
    var status = ss.getSheetByName(LOTTEON_REMOTE_TASK.statusSheet) || issue51v1Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
    if (!out || out.getLastRow() < 1) throw new Error('ISSUE51 preview 시트가 없습니다.');

    var pre = issue51v1Precheck_(ss);
    issue51v1Assert_(pre.vatRows === 3894 && pre.vatOrders === 1893 && Math.round(pre.vatPurchase) === 106707957, '실행 중 VAT 원천이 변경되었습니다.');
    issue51v1Assert_(pre.verifyOrders === 1355 && pre.overlap === 1355 && pre.currentOnly === 538 && pre.paymentReuse === 1270, '실행 중 기존검증/결제수단 기준이 변경되었습니다.');
    issue51v1Assert_(pre.verifyDup === 0 && pre.vatNormCollisions === 0, '실행 중 정규화키 충돌이 발생했습니다.');

    var processed = Math.max(0, out.getLastRow() - 1);
    var used = issue51v1UsedFromPreview_(out);
    var batch = issue51v1RunBatch_(ss, processed, ISSUE51_V1_BATCH, used, pre.paymentMap, pre.oldKeys);
    issue51v1Assert_(batch.targetOrders === 1893, 'main 카드매칭 대상 주문수 불일치: ' + batch.targetOrders);
    issue51v1Assert_(Math.round(batch.purchaseTotal) === 106707957, 'main 그룹 매입합계 불일치: ' + Math.round(batch.purchaseTotal));
    if (!batch.rows || !batch.rows.length) throw new Error('처리할 주문이 남았는데 배치 결과가 0건입니다. processed=' + processed);

    out.getRange(out.getLastRow()+1,1,batch.rows.length,issue51v1Headers_().length).setValues(batch.rows);
    var next = processed + batch.rows.length;
    SpreadsheetApp.flush();

    if (next >= batch.targetOrders) {
      var result = issue51v1Finalize_(ss, batch.canonicalRows);
      st.done = true;
      st.completedAt = new Date().toISOString();
      props.setProperty(ISSUE51_V1_STATE_KEY, JSON.stringify(st));
      issue51v1DeleteContinueTriggers_();
      return result;
    }

    issue51v1WriteStatus_(status, [
      ['항목','값'],['버전','v1.0-ISSUE51-CARD-REMATCH-PREVIEW'],['상태','RUNNING'],['단계','BATCH'],
      ['메시지','배치 처리 완료; 다음 배치 자동 예약'],['운영시트 변경','0'],
      ['처리주문',next],['대상주문',batch.targetOrders],['배치크기',ISSUE51_V1_BATCH],
      ['canonical증빙행',batch.canonicalRows],['기존결제수단재사용가능',1270],
      ['시작시각',st.startedAt],['갱신시각',new Date().toISOString()]
    ]);
    issue51v1ScheduleContinue_();
    return {ok:true,done:false,processed:next,target:batch.targetOrders,nextScheduled:true};
  } catch (e) {
    try {
      var st2 = issue51v1ReadState_(PropertiesService.getScriptProperties());
      var ss2 = st2 && st2.spreadsheetId ? SpreadsheetApp.openById(st2.spreadsheetId) : SpreadsheetApp.getActive();
      if (ss2) issue51v1Fail_(ss2, e);
    } catch (ignore) {}
    issue51v1DeleteContinueTriggers_();
    throw e;
  } finally {
    lock.releaseLock();
  }
}

function issue51v1RunBatch_(ss, startIndex, batchSize, used, paymentMap, oldKeys) {
  var code = issue51v1Fetch_(ISSUE51_V1_CODE_URL);
  var bootstrap = issue51v1Fetch_(ISSUE51_V1_BOOTSTRAP_URL);
  var sid = ss.getId();
  var invocation = [
    ';(function(){',
    'function N(v){return String(v==null?"":v).trim().toLowerCase().replace(/[^0-9a-z가-힣]/g,"");}',
    'var ss=SpreadsheetApp.openById(' + JSON.stringify(sid) + ');',
    'var detail=ss.getSheetByName("부가세_신고자료");',
    'if(!detail||detail.getLastRow()<2)throw new Error("부가세_신고자료가 없습니다.");',
    'var allOrders=groupVatDetailByOrder_v660_(detail.getDataRange().getValues());',
    'var orders=allOrders.filter(function(o){return String(o.year)==="2026"&&String(o.half)==="상반기";});',
    'orders.sort(function(a,b){return String(a.orderDate||"").localeCompare(String(b.orderDate||""))||String(a.orderNo||"").localeCompare(String(b.orderNo||""))||Number(a.purchase||0)-Number(b.purchase||0);});',
    'var pmap=' + JSON.stringify(paymentMap || {}) + ';',
    'var old=' + JSON.stringify(oldKeys || {}) + ';',
    'var purchaseTotal=0;',
    'orders.forEach(function(o){var k=String(o.account||"").trim().toLowerCase()+"|"+N(o.orderNo);o.lottePayment=pmap[k]||"";o.__issue51Old=!!old[k];purchaseTotal+=Number(o.purchase||0);});',
    'var history=loadVatCardHistory_v660_(ss);',
    'var master=loadVatCardMaster_v660_(ss);',
    'var canonical=canonicalizeVatHistory_v664_(history,master);',
    'var used=' + JSON.stringify(used || {}) + ';',
    'var start=' + Number(startIndex || 0) + ';',
    'var end=Math.min(orders.length,start+' + Number(batchSize || ISSUE51_V1_BATCH) + ');',
    'var rows=[];',
    'for(var i=start;i<end;i++){',
      'var o=orders[i];',
      'var m=matchVatOrderCardCanonical_v664_(o,canonical,master,used)||noMatch_v660_("미실행");',
      'var paymentSource=o.lottePayment?"기존검증재사용":"공란";',
      'rows.push([',
        'o.year,o.half,o.orderDate,o.business,o.account,o.orderNo,o.lottePayment,paymentSource,(o.__issue51Old?"기존1355":"신규538"),o.detailRows,o.purchase,',
        'm.company||"",m.alias||"",m.cardName||"",m.cardNumber||"",m.cardEnd4||"",',
        'm.approvalDate||"",m.approvalTime||"",m.approvalNo||"",Number(m.approvalAmount||0),',
        'm.status||"NO_MATCH",m.reason||"",Number(m.candidateCount||0),m.merchant||"",m.merchantOrderNo||"",',
        'm.evidenceType||"",m.cancelMemo||"",m.sourceFile||"",m.candidateSummary||"",',
        'm.v669Fallback?"Y":"",m.v670Fallback?"Y":"",m.canonicalEvidenceKey||""',
      ']);',
    '}',
    'return {rows:rows,targetOrders:orders.length,purchaseTotal:purchaseTotal,canonicalRows:canonical.length};',
    '})()'
  ].join('\n');
  return eval(code + '\n\n;\n\n' + bootstrap + '\n\n;\n\n' + invocation);
}

function issue51v1Finalize_(ss, canonicalRows) {
  var out = ss.getSheetByName(LOTTEON_REMOTE_TASK.outputSheet);
  var status = ss.getSheetByName(LOTTEON_REMOTE_TASK.statusSheet) || issue51v1Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  var values = out.getDataRange().getValues();
  var headers = values[0].map(issue51v1Text_);
  var ix = {};
  headers.forEach(function(h,i){ix[h]=i;});
  issue51v1Assert_(values.length-1 === 1893, 'preview 주문수 불일치: ' + (values.length-1));

  var stats = {
    orders:0, matched:0, nonCard:0, ambiguous:0, noMatch:0,
    v669:0, v669Matched:0, v669NonCard:0, v669Ambiguous:0, v670:0,
    purchase:0, invalidIdentity:0, invalidFallbackEvidence:0,
    paymentReused:0, paymentBlank:0,
    oldOrders:0,newOrders:0,
    oldMatched:0,oldNonCard:0,oldAmbiguous:0,oldNoMatch:0,
    newMatched:0,newNonCard:0,newAmbiguous:0,newNoMatch:0
  };

  for (var r=1;r<values.length;r++) {
    var row=values[r];
    var s=issue51v1Text_(row[ix['카드매칭상태']]);
    var v669=issue51v1Text_(row[ix['v6.69 2차귀속']])==='Y';
    var v670=issue51v1Text_(row[ix['v6.70 3차귀속']])==='Y';
    var origin=issue51v1Text_(row[ix['주문구분']]);
    var paymentSource=issue51v1Text_(row[ix['결제수단근거']]);
    stats.orders++;
    stats.purchase += issue51v1Num_(row[ix['주문매입금액']]);
    if (paymentSource === '기존검증재사용') stats.paymentReused++; else stats.paymentBlank++;

    if (s === 'MATCHED' || s === 'MASTER_MATCHED') stats.matched++;
    else if (s === 'NON_CARD') stats.nonCard++;
    else if (s === 'AMBIGUOUS') stats.ambiguous++;
    else stats.noMatch++;

    if (v669) {
      stats.v669++;
      if (s === 'MATCHED' || s === 'MASTER_MATCHED') stats.v669Matched++;
      else if (s === 'NON_CARD') stats.v669NonCard++;
      else if (s === 'AMBIGUOUS') stats.v669Ambiguous++;
    }
    if (v670) stats.v670++;

    if (origin === '기존1355') {
      stats.oldOrders++;
      if (s === 'MATCHED' || s === 'MASTER_MATCHED') stats.oldMatched++;
      else if (s === 'NON_CARD') stats.oldNonCard++;
      else if (s === 'AMBIGUOUS') stats.oldAmbiguous++;
      else stats.oldNoMatch++;
    } else {
      stats.newOrders++;
      if (s === 'MATCHED' || s === 'MASTER_MATCHED') stats.newMatched++;
      else if (s === 'NON_CARD') stats.newNonCard++;
      else if (s === 'AMBIGUOUS') stats.newAmbiguous++;
      else stats.newNoMatch++;
    }

    var company=issue51v1Text_(row[ix['구매카드사']]);
    var cardName=issue51v1Text_(row[ix['구매카드명']]);
    var end4=issue51v1Digits_(row[ix['카드번호끝4']]);
    if (end4) end4=('0000'+end4).slice(-4);
    var c=issue51v1Compact_(company);
    var n=issue51v1Compact_(cardName);
    if ((c.indexOf('kb')>=0 || c.indexOf('국민')>=0) && end4 && end4!=='4091') stats.invalidIdentity++;
    if (c.indexOf('우리')>=0 && end4 && end4!=='7680') stats.invalidIdentity++;
    if ((n.indexOf('tripto로카')>=0 || n.indexOf('트립투로카')>=0) && end4 && end4!=='0126') stats.invalidIdentity++;
    if ((n.indexOf('localikit')>=0 || n.indexOf('로카likit')>=0 || n.indexOf('로카리킷')>=0) && end4 && end4!=='0036') stats.invalidIdentity++;

    if (v669 || v670) {
      var approvalDate=issue51v1Text_(row[ix['승인일']]);
      var approvalNo=issue51v1Text_(row[ix['승인번호']]);
      var approvalAmount=issue51v1Num_(row[ix['승인금액']]);
      var reason=issue51v1Text_(row[ix['카드매칭근거']]);
      if (approvalDate || approvalNo || approvalAmount !== 0 || (reason.indexOf('금액비교없음') < 0 && s !== 'AMBIGUOUS')) stats.invalidFallbackEvidence++;
    }
  }

  var classified=stats.matched+stats.nonCard+stats.ambiguous+stats.noMatch;
  issue51v1Assert_(stats.orders===1893, '최종 주문수 불일치: '+stats.orders);
  issue51v1Assert_(classified===1893, '카드상태 합계 불일치: '+classified);
  issue51v1Assert_(Math.round(stats.purchase)===106707957, '주문매입금액 합계 불일치: '+Math.round(stats.purchase));
  issue51v1Assert_(stats.oldOrders===1355 && stats.newOrders===538, '기존/신규 주문 구분 불일치: '+stats.oldOrders+'/'+stats.newOrders);
  issue51v1Assert_(stats.paymentReused===1270, '결제수단 재사용 주문수 불일치: '+stats.paymentReused);
  issue51v1Assert_(stats.invalidIdentity===0, '잘못된 카드 identity: '+stats.invalidIdentity);
  issue51v1Assert_(stats.invalidFallbackEvidence===0, 'fallback 증빙필드 오류: '+stats.invalidFallbackEvidence);

  var rows=[
    ['항목','값'],['버전','v1.0-ISSUE51-CARD-REMATCH-PREVIEW'],['상태','PASS'],['단계','DONE'],
    ['메시지','현재 VAT 1,893주문 카드매칭 전체 preview 재계산 완료'],['운영시트 변경','0'],
    ['현재VAT주문',stats.orders],['preview주문',stats.orders],['canonical증빙행',canonicalRows],
    ['MATCHED',stats.matched],['NON_CARD',stats.nonCard],['AMBIGUOUS',stats.ambiguous],['NO_MATCH',stats.noMatch],
    ['v6.69 2차귀속',stats.v669],['2차귀속 MATCHED',stats.v669Matched],['2차귀속 NON_CARD',stats.v669NonCard],['2차귀속 AMBIGUOUS',stats.v669Ambiguous],
    ['v6.70 3차귀속',stats.v670],['주문매입금액합계',Math.round(stats.purchase)],
    ['기존1355주문',stats.oldOrders],['기존1355_MATCHED',stats.oldMatched],['기존1355_NON_CARD',stats.oldNonCard],['기존1355_AMBIGUOUS',stats.oldAmbiguous],['기존1355_NO_MATCH',stats.oldNoMatch],
    ['신규538주문',stats.newOrders],['신규538_MATCHED',stats.newMatched],['신규538_NON_CARD',stats.newNonCard],['신규538_AMBIGUOUS',stats.newAmbiguous],['신규538_NO_MATCH',stats.newNoMatch],
    ['기존결제수단재사용',stats.paymentReused],['결제수단공란',stats.paymentBlank],
    ['잘못된카드identity',stats.invalidIdentity],['fallback증빙필드오류',stats.invalidFallbackEvidence],
    ['부가세_카드매칭검증 변경','0'],['완료시각',new Date().toISOString()]
  ];
  issue51v1WriteStatus_(status,rows);
  try { MailApp.sendEmail('beliun1001@gmail.com','[LOTTEON 자동작업 결과][PASS] ISSUE51-v1.0',rows.map(function(x){return x[0]+': '+x[1];}).join('\n')); } catch(ignore) {}
  return {ok:true,done:true,stats:stats,canonicalRows:canonicalRows};
}

function issue51v1Precheck_(ss) {
  var vat=ss.getSheetByName('부가세_신고자료');
  var verify=ss.getSheetByName('부가세_카드매칭검증');
  var history=ss.getSheetByName('카드사용내역_붙여넣기');
  var master=ss.getSheetByName('카드_마스터');
  if(!vat||vat.getLastRow()<2)throw new Error('부가세_신고자료가 없습니다.');
  if(!verify||verify.getLastRow()<2)throw new Error('부가세_카드매칭검증이 없습니다.');
  if(!history||history.getLastRow()<2)throw new Error('카드사용내역_붙여넣기가 없습니다.');
  if(!master||master.getLastRow()<2)throw new Error('카드_마스터가 없습니다.');

  var vv=vat.getDataRange().getValues(), vh=vv[0].map(issue51v1Text_);
  var vi={year:issue51v1FindHeader_(vh,['신고연도']),half:issue51v1FindHeader_(vh,['반기']),account:issue51v1FindHeader_(vh,['쿠팡계정ID']),order:issue51v1FindHeader_(vh,['주문번호']),purchase:issue51v1FindHeader_(vh,['매입금액'])};
  issue51v1Assert_(vi.year>=0&&vi.half>=0&&vi.account>=0&&vi.order>=0&&vi.purchase>=0,'VAT 필수 헤더 누락');
  var vatMap={}, rawMap={}, vatRows=0, vatPurchase=0;
  for(var r=1;r<vv.length;r++){
    var row=vv[r]; if(issue51v1Text_(row[vi.year])!=='2026'||issue51v1Text_(row[vi.half])!=='상반기')continue;
    vatRows++; vatPurchase+=issue51v1Num_(row[vi.purchase]);
    var a=issue51v1Text_(row[vi.account]).toLowerCase(), raw=issue51v1Text_(row[vi.order]), n=issue51v1NormOrder_(raw); if(!a||!n)continue;
    var k=a+'|'+n; if(!vatMap[k])vatMap[k]=true; if(!rawMap[k])rawMap[k]={}; rawMap[k][raw]=true;
  }
  var vatKeys=Object.keys(vatMap), collisions=0; vatKeys.forEach(function(k){if(Object.keys(rawMap[k]||{}).length>1)collisions++;});

  var qv=verify.getDataRange().getValues(), qh=qv[0].map(issue51v1Text_);
  var qi={year:issue51v1FindHeader_(qh,['신고연도']),half:issue51v1FindHeader_(qh,['반기']),account:issue51v1FindHeader_(qh,['쿠팡계정ID']),order:issue51v1FindHeader_(qh,['주문번호']),payment:issue51v1FindHeader_(qh,['롯데결제수단'])};
  issue51v1Assert_(qi.year>=0&&qi.half>=0&&qi.account>=0&&qi.order>=0,'기존검증 필수 헤더 누락');
  var oldKeys={}, paymentMap={}, verifyDup=0;
  for(var i=1;i<qv.length;i++){
    var qr=qv[i]; if(issue51v1Text_(qr[qi.year])!=='2026'||issue51v1Text_(qr[qi.half])!=='상반기')continue;
    var qa=issue51v1Text_(qr[qi.account]).toLowerCase(), qo=issue51v1NormOrder_(qr[qi.order]); if(!qa||!qo)continue;
    var qk=qa+'|'+qo; if(oldKeys[qk])verifyDup++; oldKeys[qk]=true;
    var p=qi.payment>=0?issue51v1Text_(qr[qi.payment]):''; if(p)paymentMap[qk]=p;
  }
  var verifyKeys=Object.keys(oldKeys), overlap=0, currentOnly=0, paymentReuse=0;
  vatKeys.forEach(function(k){if(oldKeys[k]){overlap++;if(paymentMap[k])paymentReuse++;}else currentOnly++;});
  return {vatRows:vatRows,vatOrders:vatKeys.length,vatPurchase:vatPurchase,verifyOrders:verifyKeys.length,overlap:overlap,currentOnly:currentOnly,paymentReuse:paymentReuse,verifyDup:verifyDup,vatNormCollisions:collisions,paymentMap:paymentMap,oldKeys:oldKeys};
}

function issue51v1UsedFromPreview_(sheet) {
  var used={}; if(!sheet||sheet.getLastRow()<2)return used;
  var headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(issue51v1Text_);
  var ix=headers.indexOf('canonicalEvidenceKey'); if(ix<0)return used;
  var vals=sheet.getRange(2,ix+1,sheet.getLastRow()-1,1).getValues();
  vals.forEach(function(r){var k=issue51v1Text_(r[0]);if(k)used[k]=true;}); return used;
}
function issue51v1Headers_(){return ['신고연도','반기','주문일','사업자등록번호','쿠팡계정ID','주문번호','롯데결제수단','결제수단근거','주문구분','상세행수','주문매입금액','구매카드사','구매카드별칭','구매카드명','카드번호','카드번호끝4','승인일','승인시각','승인번호','승인금액','카드매칭상태','카드매칭근거','후보수','가맹점명','가맹점주문번호','증빙유형','취소/부분취소메모','원본파일','후보요약','v6.69 2차귀속','v6.70 3차귀속','canonicalEvidenceKey'];}
function issue51v1Fetch_(url){var res=UrlFetchApp.fetch(url+'?ts='+new Date().getTime(),{method:'get',muteHttpExceptions:true,followRedirects:true});var code=res.getResponseCode(),text=res.getContentText('UTF-8');if(code<200||code>=300)throw new Error('원격 코드 로드 실패 HTTP '+code+': '+url+'\n'+text.slice(0,500));return text;}
function issue51v1ScheduleContinue_(){issue51v1DeleteContinueTriggers_();ScriptApp.newTrigger('runLotteonRemoteTaskContinue').timeBased().after(60*1000).create();}
function issue51v1DeleteContinueTriggers_(){ScriptApp.getProjectTriggers().forEach(function(t){try{if(t.getHandlerFunction&&t.getHandlerFunction()==='runLotteonRemoteTaskContinue')ScriptApp.deleteTrigger(t);}catch(e){}});}
function issue51v1Fail_(ss,e){var props=PropertiesService.getScriptProperties();var st=issue51v1ReadState_(props)||{taskId:LOTTEON_REMOTE_TASK.id,spreadsheetId:ss?ss.getId():''};st.terminalError=true;st.error=String(e&&e.message?e.message:e);st.failedAt=new Date().toISOString();props.setProperty(ISSUE51_V1_STATE_KEY,JSON.stringify(st));var sh=ss?issue51v1Ensure_(ss,LOTTEON_REMOTE_TASK.statusSheet):null;if(sh)issue51v1WriteStatus_(sh,[['항목','값'],['버전','v1.0-ISSUE51-CARD-REMATCH-PREVIEW'],['상태','ERROR'],['단계','FAILED'],['메시지','카드매칭 전체 preview 실패'],['오류',st.error],['운영시트 변경','0'],['완료시각',st.failedAt]]);}
function issue51v1ReadState_(props){try{var s=props.getProperty(ISSUE51_V1_STATE_KEY);return s?JSON.parse(s):null;}catch(e){return null;}}
function issue51v1FindHeader_(h,names){for(var n=0;n<names.length;n++){var w=issue51v1Compact_(names[n]);for(var i=0;i<h.length;i++)if(issue51v1Compact_(h[i])===w)return i;}return -1;}
function issue51v1Text_(v){return String(v==null?'':v).trim();}
function issue51v1Compact_(v){return issue51v1Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function issue51v1NormOrder_(v){return issue51v1Text_(v).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');}
function issue51v1Digits_(v){return issue51v1Text_(v).replace(/\D/g,'');}
function issue51v1Num_(v){if(typeof v==='number'&&isFinite(v))return v;var n=Number(issue51v1Text_(v).replace(/[원,%\s,]/g,''));return isFinite(n)?n:0;}
function issue51v1Assert_(ok,msg){if(!ok)throw new Error(msg);}
function issue51v1Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}
function issue51v1WriteStatus_(sh,rows){sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setFrozenRows(1);}
