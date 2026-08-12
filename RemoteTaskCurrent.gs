/**
 * Issue #46 standalone diagnostic.
 * Compares source/VAT order-number raw/display/type/format for the 5 normalized matches.
 * Reads production sheets; writes only ISSUE46_* diagnostic sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE46-v1.0-20260812',
  title: '상반기 VAT 주문번호 표시형식 불일치 5건 원문·셀타입 진단',
  enabled: true,
  inputSheet: 'ISSUE45_VAT상세누락추적',
  outputSheet: 'ISSUE46_주문번호표시형식진단',
  statusSheet: 'ISSUE46_진단상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var state = issue46Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue46Write_(state,[
    ['항목','값'],['버전','v1.0-ISSUE46-ORDER-FORMAT-DIAGNOSTIC'],['상태','RUNNING'],['단계','LOAD'],
    ['메시지','주문번호 표시형식 5건 원문·셀타입 진단 시작'],['운영시트 변경','0'],['갱신시각',new Date().toISOString()]
  ]);

  try {
    var input=ss.getSheetByName(LOTTEON_REMOTE_TASK.inputSheet);
    var source=ss.getSheetByName('매출데이터_붙여넣기');
    var vat=ss.getSheetByName('부가세_신고자료');
    if(!input||input.getLastRow()<2) throw new Error('ISSUE45_VAT상세누락추적 시트가 없습니다.');
    if(!source||source.getLastRow()<2) throw new Error('매출데이터_붙여넣기 시트가 없습니다.');
    if(!vat||vat.getLastRow()<2) throw new Error('부가세_신고자료 시트가 없습니다.');

    var iv=input.getDataRange().getValues(), ih=issue46Map_(iv[0]);
    var io=issue46Find_(ih,['주문번호']);
    var ia=issue46Find_(ih,['쿠팡계정ID']);
    if(io<0||ia<0) throw new Error('ISSUE45 필수 헤더 누락');
    var targets=[];
    for(var r=1;r<iv.length;r++){
      var o=issue46Text_(iv[r][io]);
      if(!o) continue;
      targets.push({order:o,account:issue46Text_(iv[r][ia]),issue45Row:r+1});
    }
    if(targets.length!==5) throw new Error('대상 건수 불일치: 기대 5건, 실제 '+targets.length+'건');

    var sr=source.getDataRange();
    var sv=sr.getValues(), sd=sr.getDisplayValues(), sf=sr.getNumberFormats(), sform=sr.getFormulas();
    var sh=sv[0]||[], sm=issue46Map_(sh);
    var so=issue46Find_(sm,['마켓주문번호','주문번호','주문ID','주문ID(마켓)']);
    var sa=issue46Find_(sm,['마켓아이디','쿠팡계정ID','계정ID']);
    if(so<0) throw new Error('매출데이터_붙여넣기 주문번호 헤더를 찾지 못했습니다.');

    var vr=vat.getDataRange();
    var vv=vr.getValues(), vd=vr.getDisplayValues(), vf=vr.getNumberFormats(), vform=vr.getFormulas();
    var vh=vv[0]||[], vm=issue46Map_(vh);
    var vo=issue46Find_(vm,['주문번호','마켓주문번호','주문ID','주문ID(마켓)']);
    var va=issue46Find_(vm,['쿠팡계정ID','마켓아이디','계정ID']);
    if(vo<0) throw new Error('부가세_신고자료 주문번호 헤더를 찾지 못했습니다.');

    var headers=[
      'Issue45행','쿠팡계정ID','기준주문번호','기준정규화',
      '원천행번호','원천raw','원천display','원천type','원천numberFormat','원천formula','원천길이','원천비영숫자',
      'VAT행번호','VATraw','VATdisplay','VATtype','VATnumberFormat','VATformula','VAT길이','VAT비영숫자',
      '정규화동일','계정일치','차이분류','진단메모'
    ];
    var out=[], counts={};

    targets.forEach(function(t){
      var norm=issue46OrderNorm_(t.order);
      var sRows=issue46FindNormRows_(sv,so,norm);
      if(sa>=0&&t.account){
        var sAcct=sRows.filter(function(x){return issue46Text_(x.row[sa]).toLowerCase()===t.account.toLowerCase();});
        if(sAcct.length) sRows=sAcct;
      }
      var vRows=issue46FindNormRows_(vv,vo,norm);
      if(va>=0&&t.account){
        var vAcct=vRows.filter(function(x){return issue46Text_(x.row[va]).toLowerCase()===t.account.toLowerCase();});
        if(vAcct.length) vRows=vAcct;
      }
      if(!sRows.length) throw new Error('원천 정규화 주문번호 행 없음: '+t.order);
      if(!vRows.length) throw new Error('VAT 정규화 주문번호 행 없음: '+t.order);

      var s=sRows[0], v=vRows[0];
      var sRaw=sv[s.rowNo-1][so], sDisp=sd[s.rowNo-1][so], sFmt=sf[s.rowNo-1][so], sFormula=sform[s.rowNo-1][so];
      var vRaw=vv[v.rowNo-1][vo], vDisp=vd[v.rowNo-1][vo], vFmt=vf[v.rowNo-1][vo], vFormula=vform[v.rowNo-1][vo];
      var sText=issue46Text_(sRaw), vText=issue46Text_(vRaw);
      var sameNorm=issue46OrderNorm_(sText)===issue46OrderNorm_(vText);
      var accountSame=true;
      if(sa>=0&&va>=0&&t.account){
        accountSame=issue46Text_(sv[s.rowNo-1][sa]).toLowerCase()===issue46Text_(vv[v.rowNo-1][va]).toLowerCase();
      }
      var classification=issue46Class_(sText,vText,sDisp,vDisp,typeof sRaw,typeof vRaw,sFmt,vFmt);
      counts[classification]=(counts[classification]||0)+1;
      var memo=[];
      memo.push('sourceRows='+sRows.map(function(x){return x.rowNo;}).join(','));
      memo.push('vatRows='+vRows.map(function(x){return x.rowNo;}).join(','));
      memo.push('rawEqual='+(sText===vText?'Y':'N'));
      memo.push('displayEqual='+(String(sDisp)===String(vDisp)?'Y':'N'));

      out.push([
        t.issue45Row,t.account,t.order,norm,
        s.rowNo,sText,String(sDisp),typeof sRaw,sFmt,sFormula||'',sText.length,issue46Punct_(sText),
        v.rowNo,vText,String(vDisp),typeof vRaw,vFmt,vFormula||'',vText.length,issue46Punct_(vText),
        sameNorm?'Y':'N',accountSame?'Y':'N',classification,memo.join(' / ')
      ]);
    });

    if(out.length!==5) throw new Error('출력 건수 불일치: '+out.length);
    var output=issue46Ensure_(ss,LOTTEON_REMOTE_TASK.outputSheet);
    output.clearContents();
    output.getRange(1,1,1,headers.length).setValues([headers]);
    output.getRange(2,1,out.length,headers.length).setValues(out);
    output.setFrozenRows(1);
    output.getRange(1,1,1,headers.length).setBackground('#d9eaf7').setFontWeight('bold');

    var status=[
      ['항목','값'],['버전','v1.0-ISSUE46-ORDER-FORMAT-DIAGNOSTIC'],['상태','PASS'],['단계','DONE'],
      ['메시지','주문번호 표시형식 5건 원문·셀타입 진단 완료'],['대상건수',5],['출력건수',out.length],['운영시트 변경','0']
    ];
    Object.keys(counts).sort().forEach(function(k){status.push(['분류_'+k,counts[k]]);});
    status.push(['완료시각',new Date().toISOString()]);
    issue46Write_(state,status);
    try{MailApp.sendEmail('beliun1001@gmail.com','[LOTTEON 자동작업 결과][PASS] ISSUE46-v1.0',status.map(function(x){return x[0]+': '+x[1];}).join('\n'));}catch(mailError){}
    return {ok:true,counts:counts};
  } catch(e) {
    issue46Write_(state,[['항목','값'],['버전','v1.0-ISSUE46-ORDER-FORMAT-DIAGNOSTIC'],['상태','ERROR'],['단계','FAILED'],['메시지','주문번호 표시형식 5건 진단 실패'],['오류',String(e&&e.message?e.message:e)],['운영시트 변경','0'],['갱신시각',new Date().toISOString()]]);
    throw e;
  }
}

function issue46Class_(s,v,sd,vd,st,vt,sf,vf){
  if(s===v){
    if(String(sd)!==String(vd)) return 'display-only 차이';
    if(st!==vt||sf!==vf) return '셀 타입/number format 차이';
    return '기타';
  }
  if(s.replace(/\s/g,'')===v.replace(/\s/g,'')) return '공백 차이';
  var dash=/[-\u2010\u2011\u2012\u2013\u2014\u2015\u2212]/g;
  if(s.replace(dash,'')===v.replace(dash,'')) return '하이픈/구분자 차이';
  if(issue46OrderNorm_(s)===issue46OrderNorm_(v)) return '기타 문장부호 차이';
  if(st!==vt||sf!==vf) return '셀 타입/number format 차이';
  return '기타';
}
function issue46FindNormRows_(values,ix,norm){var a=[];for(var r=1;r<values.length;r++){if(issue46OrderNorm_(values[r][ix])===norm)a.push({rowNo:r+1,row:values[r]});}return a;}
function issue46Punct_(s){var a=[];for(var i=0;i<s.length;i++){var c=s.charAt(i);if(!/[0-9a-zA-Z가-힣]/.test(c)){a.push(c+'(U+'+('0000'+c.charCodeAt(0).toString(16).toUpperCase()).slice(-4)+')');}}return a.join(' ');}
function issue46OrderNorm_(v){return issue46Text_(v).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');}
function issue46Find_(m,a){for(var i=0;i<a.length;i++){var k=issue46Norm_(a[i]);if(m[k]!=null)return m[k];}return -1;}
function issue46Map_(h){var m={};(h||[]).forEach(function(x,i){m[issue46Norm_(x)]=i;});return m;}
function issue46Norm_(v){return issue46Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function issue46Text_(v){return v==null?'':String(v).trim();}
function issue46Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}
function issue46Write_(s,r){s.clearContents();s.getRange(1,1,r.length,2).setValues(r);s.setFrozenRows(1);s.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');SpreadsheetApp.flush();}
function runLotteonRemoteTaskContinueRemote_(){return runLotteonRemoteTaskStartRemote_();}
