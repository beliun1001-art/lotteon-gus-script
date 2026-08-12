/**
 * Issue #45 standalone diagnostic.
 * Traces why 5 source orders are missing from VAT detail.
 * Reads production sheets; writes only ISSUE45_* diagnostic sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE45-v1.0-20260812',
  title: '상반기 VAT 상세행 누락 5건 생성 제외조건 진단',
  enabled: true,
  inputSheet: 'ISSUE44_매입금액0원추적',
  outputSheet: 'ISSUE45_VAT상세누락추적',
  statusSheet: 'ISSUE45_진단상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var state = issue45Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue45Write_(state,[
    ['항목','값'],['버전','v1.0-ISSUE45-VAT-DETAIL-EXCLUSION-DIAGNOSTIC'],['상태','RUNNING'],['단계','LOAD'],
    ['메시지','VAT 상세행 누락 5건 생성 제외조건 진단 시작'],['운영시트 변경','0'],['갱신시각',new Date().toISOString()]
  ]);
  try {
    var input=ss.getSheetByName(LOTTEON_REMOTE_TASK.inputSheet);
    var source=ss.getSheetByName('매출데이터_붙여넣기');
    var vat=ss.getSheetByName('부가세_신고자료');
    if(!input||input.getLastRow()<2) throw new Error('ISSUE44_매입금액0원추적 시트가 없습니다.');
    if(!source||source.getLastRow()<2) throw new Error('매출데이터_붙여넣기 시트가 없습니다.');
    if(!vat||vat.getLastRow()<2) throw new Error('부가세_신고자료 시트가 없습니다.');

    var iv=input.getDataRange().getValues(), ih=issue45Map_(iv[0]);
    var orderKey=issue45Norm_('주문번호'), accountKey=issue45Norm_('쿠팡계정ID');
    if(ih[orderKey]==null||ih[accountKey]==null) throw new Error('ISSUE44 필수 헤더 누락');
    var targets=[];
    for(var r=1;r<iv.length;r++){
      var o=issue45Text_(iv[r][ih[orderKey]]); if(!o) continue;
      targets.push({order:o,account:issue45Text_(iv[r][ih[accountKey]]),issue44Row:r+1});
    }
    if(targets.length!==5) throw new Error('대상 건수 불일치: 기대 5건, 실제 '+targets.length+'건');

    var sv=source.getDataRange().getValues(), sh=sv[0]||[], sm=issue45Map_(sh);
    var vv=vat.getDataRange().getValues(), vh=vv[0]||[], vm=issue45Map_(vh);
    var so=issue45Find_(sm,['마켓주문번호','주문번호','주문ID','주문ID(마켓)']);
    var sa=issue45Find_(sm,['마켓아이디','쿠팡계정ID','계정ID']);
    var vo=issue45Find_(vm,['주문번호','마켓주문번호','주문ID','주문ID(마켓)']);
    var va=issue45Find_(vm,['쿠팡계정ID','마켓아이디','계정ID']);
    if(so<0||vo<0) throw new Error('주문번호 헤더를 찾지 못했습니다.');

    var dateIxs=issue45MatchingIndexes_(sh,/(주문.*일|결제.*일|매출.*일|등록.*일|배송.*일|일자|일시|date|time)/i);
    var statusIxs=issue45MatchingIndexes_(sh,/(상태|status|클레임|취소|반품|환불)/i);
    var marketIxs=issue45MatchingIndexes_(sh,/(마켓|판매처|쇼핑몰|채널|mall|market)/i);
    var amountIxs=issue45MatchingIndexes_(sh,/(매출|정산|수수료|매입|구매가격|판매금액|결제금액|금액)/i);

    var out=[], counts={};
    targets.forEach(function(t){
      var sourceOrder=issue45RowsExact_(sv,so,t.order);
      var sourceExact=(sa>=0&&t.account)?sourceOrder.filter(function(e){return issue45Text_(e.row[sa]).toLowerCase()===t.account.toLowerCase();}):sourceOrder.slice();
      var sourceRows=sourceExact.length?sourceExact:sourceOrder;
      var vatExact=issue45RowsExact_(vv,vo,t.order);
      var vatNorm=issue45RowsNorm_(vv,vo,t.order);
      var vatSameAccount=(va>=0&&t.account)?vatNorm.filter(function(e){return issue45Text_(e.row[va]).toLowerCase()===t.account.toLowerCase();}):vatNorm.slice();

      var dateInfo=issue45CollectDates_(sourceRows,sh,dateIxs);
      var statusText=issue45Collect_(sourceRows,sh,statusIxs);
      var marketText=issue45Collect_(sourceRows,sh,marketIxs);
      var amountText=issue45Collect_(sourceRows,sh,amountIxs);
      var accountSet={}; sourceOrder.forEach(function(e){if(sa>=0){var a=issue45Text_(e.row[sa]);if(a)accountSet[a]=true;}});
      var multiAccount=Object.keys(accountSet).length>1;
      var normalizedVatOnly=vatExact.length===0&&vatNorm.length>0;
      var hasCancel=/(취소|반품|환불|교환)/.test(statusText);
      var className=''; var memo=[];

      if(normalizedVatOnly){className='주문번호 표시형식 불일치 의심';memo.push('정규화 주문번호로 VAT행 발견='+vatNorm.length);}
      else if(sa>=0&&t.account&&sourceOrder.length>0&&sourceExact.length===0){className='계정 매칭 불일치 의심';memo.push('원천 주문번호 행은 있으나 대상 계정과 불일치');}
      else if(dateInfo.count>0&&!dateInfo.h1){className='상반기 날짜 범위 밖 의심';}
      else if(hasCancel){className='취소/반품/환불 상태 제외 의심';}
      else if(sourceOrder.length>1||multiAccount){className='동일 주문 다중행/중복 제외 의심';memo.push('원천행='+sourceOrder.length+', 계정수='+Object.keys(accountSet).length);}
      else if(sourceRows.length>0&&dateInfo.h1&&vatNorm.length===0){className='기본 적격조건은 충족하나 VAT 상세 누락';}
      else {className='기타 추가 추적 필요';}
      counts[className]=(counts[className]||0)+1;

      out.push([
        t.issue44Row,t.account,t.order,issue45OrderNorm_(t.order),sourceOrder.length,sourceExact.length,Object.keys(accountSet).join(' | '),multiAccount?'Y':'N',
        dateInfo.count,dateInfo.h1?'Y':'N',dateInfo.summary,statusText,marketText,amountText,
        vatExact.length,vatNorm.length,vatSameAccount.length,issue45RowsNo_(sourceRows),issue45RowsNo_(vatNorm),className,memo.join(' / ')
      ]);
    });

    var headers=['Issue44행','쿠팡계정ID','주문번호','주문번호정규화','원천주문행수','원천계정일치행수','원천계정값','다중계정','날짜후보수','상반기포함','날짜상세','상태상세','마켓상세','금액상세','VAT exact행수','VAT 정규화행수','VAT 정규화+계정일치','원천행번호','VAT행번호','원인후보분류','진단메모'];
    var output=issue45Ensure_(ss,LOTTEON_REMOTE_TASK.outputSheet);
    output.clearContents(); output.getRange(1,1,1,headers.length).setValues([headers]); output.getRange(2,1,out.length,headers.length).setValues(out);
    output.setFrozenRows(1); output.getRange(1,1,1,headers.length).setBackground('#d9eaf7').setFontWeight('bold');

    var status=[['항목','값'],['버전','v1.0-ISSUE45-VAT-DETAIL-EXCLUSION-DIAGNOSTIC'],['상태','PASS'],['단계','DONE'],['메시지','VAT 상세행 누락 5건 생성 제외조건 진단 완료'],['대상건수',5],['출력건수',out.length],['운영시트 변경','0']];
    Object.keys(counts).sort().forEach(function(k){status.push(['분류_'+k,counts[k]]);});
    status.push(['완료시각',new Date().toISOString()]); issue45Write_(state,status);
    try{MailApp.sendEmail('beliun1001@gmail.com','[LOTTEON 자동작업 결과][PASS] ISSUE45-v1.0',status.map(function(x){return x[0]+': '+x[1];}).join('\n'));}catch(mailError){}
    return {ok:true,counts:counts};
  } catch(e) {
    issue45Write_(state,[['항목','값'],['버전','v1.0-ISSUE45-VAT-DETAIL-EXCLUSION-DIAGNOSTIC'],['상태','ERROR'],['단계','FAILED'],['메시지','VAT 상세행 누락 5건 생성 제외조건 진단 실패'],['오류',String(e&&e.message?e.message:e)],['운영시트 변경','0'],['갱신시각',new Date().toISOString()]]);
    throw e;
  }
}

function issue45RowsExact_(v,ix,o){var a=[];for(var r=1;r<v.length;r++)if(issue45Text_(v[r][ix])===o)a.push({rowNo:r+1,row:v[r]});return a;}
function issue45RowsNorm_(v,ix,o){var n=issue45OrderNorm_(o),a=[];for(var r=1;r<v.length;r++)if(issue45OrderNorm_(v[r][ix])===n)a.push({rowNo:r+1,row:v[r]});return a;}
function issue45OrderNorm_(v){return issue45Text_(v).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');}
function issue45MatchingIndexes_(h,re){var a=[];(h||[]).forEach(function(x,i){if(re.test(issue45Text_(x)))a.push(i);});return a;}
function issue45Collect_(rows,h,ixs){var a=[];(rows||[]).forEach(function(e){(ixs||[]).forEach(function(ix){var v=issue45Text_(e.row[ix]);if(v)a.push('R'+e.rowNo+':'+issue45Text_(h[ix])+'='+v);});});return a.join(' | ');}
function issue45CollectDates_(rows,h,ixs){var a=[],h1=false;(rows||[]).forEach(function(e){(ixs||[]).forEach(function(ix){var raw=e.row[ix],d=issue45Date_(raw);if(d){if(d>=new Date(2026,0,1)&&d<new Date(2026,6,1))h1=true;a.push('R'+e.rowNo+':'+issue45Text_(h[ix])+'='+Utilities.formatDate(d,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd'));}});});return {count:a.length,h1:h1,summary:a.join(' | ')};}
function issue45Date_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return v;var s=issue45Text_(v),m=s.match(/(20\d{2})[^0-9]?(\d{1,2})[^0-9]?(\d{1,2})/);if(!m)return null;var d=new Date(+m[1],+m[2]-1,+m[3]);return isNaN(d.getTime())?null:d;}
function issue45RowsNo_(rows){return (rows||[]).map(function(e){return e.rowNo;}).join(',');}
function issue45Find_(m,a){for(var i=0;i<a.length;i++){var k=issue45Norm_(a[i]);if(m[k]!=null)return m[k];}return -1;}
function issue45Map_(h){var m={};(h||[]).forEach(function(x,i){m[issue45Norm_(x)]=i;});return m;}
function issue45Norm_(v){return issue45Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function issue45Text_(v){return v==null?'':String(v).trim();}
function issue45Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}
function issue45Write_(s,r){s.clearContents();s.getRange(1,1,r.length,2).setValues(r);s.setFrozenRows(1);s.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');SpreadsheetApp.flush();}
function runLotteonRemoteTaskContinueRemote_(){return runLotteonRemoteTaskStartRemote_();}
