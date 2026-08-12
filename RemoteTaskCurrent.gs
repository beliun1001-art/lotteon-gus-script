/**
 * Issue #47 standalone diagnostic.
 * Traces 5 normalized orders across source -> VAT detail -> card verification.
 * Reads production sheets; writes only ISSUE47_* diagnostic sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE47-v1.0-20260812',
  title: '상반기 VAT 주문번호 문장부호 5건 3단계 JOIN 경로 진단',
  enabled: true,
  inputSheet: 'ISSUE46_주문번호표시형식진단',
  outputSheet: 'ISSUE47_3단계JOIN진단',
  statusSheet: 'ISSUE47_진단상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var state = issue47Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue47Write_(state,[
    ['항목','값'],['버전','v1.0-ISSUE47-THREE-STAGE-JOIN-TRACE'],['상태','RUNNING'],['단계','LOAD'],
    ['메시지','3단계 주문번호 JOIN 경로 진단 시작'],['운영시트 변경','0'],['갱신시각',new Date().toISOString()]
  ]);

  try {
    var input = ss.getSheetByName(LOTTEON_REMOTE_TASK.inputSheet);
    var source = ss.getSheetByName('매출데이터_붙여넣기');
    var vat = ss.getSheetByName('부가세_신고자료');
    var verify = ss.getSheetByName('부가세_카드매칭검증');
    if (!input || input.getLastRow() < 2) throw new Error('ISSUE46_주문번호표시형식진단 시트가 없습니다.');
    if (!source || source.getLastRow() < 2) throw new Error('매출데이터_붙여넣기 시트가 없습니다.');
    if (!vat || vat.getLastRow() < 2) throw new Error('부가세_신고자료 시트가 없습니다.');
    if (!verify || verify.getLastRow() < 2) throw new Error('부가세_카드매칭검증 시트가 없습니다.');

    var iv=input.getDataRange().getValues(), ih=issue47Map_(iv[0]);
    var iNorm=issue47Find_(ih,['기준정규화']);
    var iAccount=issue47Find_(ih,['쿠팡계정ID']);
    var iSourceRaw=issue47Find_(ih,['원천raw']);
    var iVatRaw=issue47Find_(ih,['VATraw']);
    if(iNorm<0||iAccount<0||iSourceRaw<0||iVatRaw<0) throw new Error('ISSUE46 필수 헤더 누락');
    var targets=[];
    for(var r=1;r<iv.length;r++){
      var norm=issue47Text_(iv[r][iNorm]);
      if(!norm) continue;
      targets.push({
        norm:norm,
        account:issue47Text_(iv[r][iAccount]),
        sourceRaw:issue47Text_(iv[r][iSourceRaw]),
        vatRaw:issue47Text_(iv[r][iVatRaw]),
        issue46Row:r+1
      });
    }
    if(targets.length!==5) throw new Error('대상 건수 불일치: 기대 5건, 실제 '+targets.length+'건');

    var sv=source.getDataRange().getValues(), sh=sv[0]||[], sm=issue47Map_(sh);
    var so=issue47Find_(sm,['마켓주문번호','주문번호','주문ID','주문ID(마켓)']);
    var sa=issue47Find_(sm,['마켓아이디','쿠팡계정ID','계정ID']);
    if(so<0) throw new Error('매출데이터_붙여넣기 주문번호 헤더를 찾지 못했습니다.');
    var sourceAcIx=sh.length>28?28:-1;
    if(sourceAcIx<0) throw new Error('매출데이터_붙여넣기 AC열(29열)이 없습니다.');
    var sourcePurchaseIxs=issue47FindAll_(sh,['매입금액','구매가격','매입가격','상품매입금액','결제금액(매입)']);

    var vv=vat.getDataRange().getValues(), vh=vv[0]||[], vm=issue47Map_(vh);
    var vo=issue47Find_(vm,['주문번호','마켓주문번호','주문ID','주문ID(마켓)']);
    var va=issue47Find_(vm,['쿠팡계정ID','마켓아이디','계정ID']);
    if(vo<0) throw new Error('부가세_신고자료 주문번호 헤더를 찾지 못했습니다.');
    var vatPurchaseIxs=issue47FindAll_(vh,['주문매입금액','매입금액','구매가격','매입가격','상품매입금액']);
    if(!vatPurchaseIxs.length) throw new Error('부가세_신고자료 매입금액 헤더를 찾지 못했습니다.');

    var cv=verify.getDataRange().getValues(), ch=cv[0]||[], cm=issue47Map_(ch);
    var co=issue47Find_(cm,['주문번호','마켓주문번호','주문ID','주문ID(마켓)']);
    var ca=issue47Find_(cm,['쿠팡계정ID','마켓아이디','계정ID']);
    var cp=issue47Find_(cm,['주문매입금액','매입금액']);
    var cs=issue47Find_(cm,['카드매칭상태']);
    if(co<0||cp<0) throw new Error('부가세_카드매칭검증 주문번호/주문매입금액 헤더를 찾지 못했습니다.');

    var headers=[
      'Issue46행','쿠팡계정ID','정규화주문번호',
      '원천주문번호','원천문장부호','원천행수','원천AC합계','원천매입별칭합계','원천행번호',
      'VAT주문번호','VAT문장부호','VAT행수','VAT매입금액합계','VAT행번호',
      '검증주문번호','검증문장부호','검증행수','검증주문매입금액합계','검증카드상태','검증행번호',
      '검증표현관계','정규화3단계동일','원천AC양수','VAT매입양수','검증금액0','원인분류','진단메모'
    ];
    var out=[], counts={}, relationCounts={}, patternCounts={};
    var sourceAcPositive=0, vatPositive=0, verifyZero=0, allNormSame=0, vatPosVerifyZero=0;

    targets.forEach(function(t){
      var sRows=issue47RowsNorm_(sv,so,t.norm);
      sRows=issue47PreferAccount_(sRows,sa,t.account);
      var vRows=issue47RowsNorm_(vv,vo,t.norm);
      vRows=issue47PreferAccount_(vRows,va,t.account);
      var cRows=issue47RowsNorm_(cv,co,t.norm);
      cRows=issue47PreferAccount_(cRows,ca,t.account);
      if(!sRows.length) throw new Error('원천 정규화 주문행 없음: '+t.norm);
      if(!vRows.length) throw new Error('VAT 정규화 주문행 없음: '+t.norm);
      if(!cRows.length) throw new Error('카드검증 정규화 주문행 없음: '+t.norm);

      var sourceRaw=issue47Text_(sRows[0].row[so]);
      var vatRaw=issue47Text_(vRows[0].row[vo]);
      var verifyRaw=issue47Text_(cRows[0].row[co]);
      var sourceAc=issue47SumIx_(sRows,[sourceAcIx]);
      var sourceNamed=issue47SumIx_(sRows,sourcePurchaseIxs.filter(function(ix){return ix!==sourceAcIx;}));
      var vatPurchase=issue47SumIx_(vRows,vatPurchaseIxs);
      var verifyPurchase=issue47SumIx_(cRows,[cp]);
      var statuses=cs>=0?issue47Distinct_(cRows.map(function(e){return issue47Text_(e.row[cs]);})).join(' | '):'';

      var relation='제3표현';
      if(verifyRaw===sourceRaw && verifyRaw===vatRaw) relation='원천=VAT=검증 동일';
      else if(verifyRaw===sourceRaw) relation='원천표현동일';
      else if(verifyRaw===vatRaw) relation='VAT표현동일';
      relationCounts[relation]=(relationCounts[relation]||0)+1;

      var normSame=issue47OrderNorm_(sourceRaw)===t.norm && issue47OrderNorm_(vatRaw)===t.norm && issue47OrderNorm_(verifyRaw)===t.norm;
      var acPos=sourceAc>0, vatPos=vatPurchase>0, verZero=verifyPurchase===0;
      if(acPos) sourceAcPositive++;
      if(vatPos) vatPositive++;
      if(verZero) verifyZero++;
      if(normSame) allNormSame++;
      if(vatPos&&verZero) vatPosVerifyZero++;

      var classification='기타 추가추적';
      if(vatPurchase===0) classification='VAT매입 자체 0';
      else if(verifyPurchase>0) classification='검증금액 정상';
      else if(vatPos&&verZero&&relation==='원천표현동일') classification='검증=원천표현 / VAT양수 / 검증0';
      else if(vatPos&&verZero&&relation==='VAT표현동일') classification='검증=VAT표현 / VAT양수 / 검증0';
      else if(vatPos&&verZero) classification='검증=제3표현 / VAT양수 / 검증0';
      counts[classification]=(counts[classification]||0)+1;

      var sourceP=issue47Punct_(sourceRaw), vatP=issue47Punct_(vatRaw), verifyP=issue47Punct_(verifyRaw);
      var pattern='원천['+(sourceP||'없음')+'] → VAT['+(vatP||'없음')+'] → 검증['+(verifyP||'없음')+']';
      patternCounts[pattern]=(patternCounts[pattern]||0)+1;
      var memo=[
        'sourceRawEqIssue46='+(sourceRaw===t.sourceRaw?'Y':'N'),
        'vatRawEqIssue46='+(vatRaw===t.vatRaw?'Y':'N'),
        'sourceAC='+sourceAc,
        'vatPurchase='+vatPurchase,
        'verifyPurchase='+verifyPurchase
      ];

      out.push([
        t.issue46Row,t.account,t.norm,
        sourceRaw,sourceP,sRows.length,sourceAc,sourceNamed,issue47RowNos_(sRows),
        vatRaw,vatP,vRows.length,vatPurchase,issue47RowNos_(vRows),
        verifyRaw,verifyP,cRows.length,verifyPurchase,statuses,issue47RowNos_(cRows),
        relation,normSame?'Y':'N',acPos?'Y':'N',vatPos?'Y':'N',verZero?'Y':'N',classification,memo.join(' / ')
      ]);
    });

    if(out.length!==5) throw new Error('출력 건수 불일치: '+out.length);
    var output=issue47Ensure_(ss,LOTTEON_REMOTE_TASK.outputSheet);
    output.clearContents();
    output.getRange(1,1,1,headers.length).setValues([headers]);
    output.getRange(2,1,out.length,headers.length).setValues(out);
    output.setFrozenRows(1);
    output.getRange(1,1,1,headers.length).setBackground('#d9eaf7').setFontWeight('bold');
    [7,8,13,18].forEach(function(c){output.getRange(2,c,out.length,1).setNumberFormat('#,##0');});

    var status=[
      ['항목','값'],['버전','v1.0-ISSUE47-THREE-STAGE-JOIN-TRACE'],['상태','PASS'],['단계','DONE'],
      ['메시지','3단계 주문번호 JOIN 경로 진단 완료'],['대상건수',5],['출력건수',out.length],['운영시트 변경','0'],
      ['원천AC양수',sourceAcPositive],['VAT매입양수',vatPositive],['검증금액0',verifyZero],['VAT양수_검증0',vatPosVerifyZero],['정규화3단계동일',allNormSame]
    ];
    Object.keys(relationCounts).sort().forEach(function(k){status.push(['검증주문번호_'+k,relationCounts[k]]);});
    Object.keys(counts).sort().forEach(function(k){status.push(['분류_'+k,counts[k]]);});
    Object.keys(patternCounts).sort().forEach(function(k,i){status.push(['문장부호패턴_'+(i+1),k+' / '+patternCounts[k]+'건']);});
    status.push(['완료시각',new Date().toISOString()]);
    issue47Write_(state,status);
    try{MailApp.sendEmail('beliun1001@gmail.com','[LOTTEON 자동작업 결과][PASS] ISSUE47-v1.0',status.map(function(x){return x[0]+': '+x[1];}).join('\n'));}catch(mailError){}
    return {ok:true,counts:counts,relations:relationCounts};
  } catch(e) {
    issue47Write_(state,[
      ['항목','값'],['버전','v1.0-ISSUE47-THREE-STAGE-JOIN-TRACE'],['상태','ERROR'],['단계','FAILED'],
      ['메시지','3단계 주문번호 JOIN 경로 진단 실패'],['오류',String(e&&e.message?e.message:e)],['운영시트 변경','0'],['갱신시각',new Date().toISOString()]
    ]);
    throw e;
  }
}

function issue47RowsNorm_(values,ix,norm){var a=[];for(var r=1;r<values.length;r++){if(issue47OrderNorm_(values[r][ix])===norm)a.push({rowNo:r+1,row:values[r]});}return a;}
function issue47PreferAccount_(rows,ix,account){if(ix<0||!account)return rows;var a=rows.filter(function(e){return issue47Text_(e.row[ix]).toLowerCase()===account.toLowerCase();});return a.length?a:rows;}
function issue47SumIx_(rows,ixs){var n=0;(rows||[]).forEach(function(e){(ixs||[]).forEach(function(ix){n+=issue47Number_(e.row[ix]);});});return n;}
function issue47FindAll_(headers,aliases){var set={};aliases.forEach(function(a){set[issue47Norm_(a)]=true;});var out=[];(headers||[]).forEach(function(h,i){if(set[issue47Norm_(h)])out.push(i);});return out;}
function issue47Distinct_(a){var m={},o=[];(a||[]).forEach(function(v){if(v&&!m[v]){m[v]=true;o.push(v);}});return o;}
function issue47Punct_(s){var a=[];s=issue47Text_(s);for(var i=0;i<s.length;i++){var c=s.charAt(i);if(!/[0-9a-zA-Z가-힣]/.test(c)){var x=c+'(U+'+('0000'+c.charCodeAt(0).toString(16).toUpperCase()).slice(-4)+')';if(a.indexOf(x)<0)a.push(x);}}return a.join(' ');}
function issue47OrderNorm_(v){return issue47Text_(v).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');}
function issue47RowNos_(rows){return (rows||[]).map(function(e){return e.rowNo;}).join(',');}
function issue47Find_(m,a){for(var i=0;i<a.length;i++){var k=issue47Norm_(a[i]);if(m[k]!=null)return m[k];}return -1;}
function issue47Map_(h){var m={};(h||[]).forEach(function(x,i){m[issue47Norm_(x)]=i;});return m;}
function issue47Norm_(v){return issue47Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function issue47Text_(v){return v==null?'':String(v).trim();}
function issue47Number_(v){if(typeof v==='number'&&isFinite(v))return v;var n=Number(String(v==null?'':v).replace(/,/g,'').replace(/[^0-9.\-]/g,''));return isFinite(n)?n:0;}
function issue47Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}
function issue47Write_(s,r){s.clearContents();s.getRange(1,1,r.length,2).setValues(r);s.setFrozenRows(1);s.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');SpreadsheetApp.flush();}
function runLotteonRemoteTaskContinueRemote_(){return runLotteonRemoteTaskStartRemote_();}
