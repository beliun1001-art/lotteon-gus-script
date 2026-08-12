/**
 * Issue #44 standalone diagnostic.
 * Traces the 5 orders where source purchase amount exists but VAT detail is 0.
 * Reads production sheets; writes only ISSUE44_* diagnostic sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE44-v1.0-20260812',
  title: '상반기 VAT 매입금액 0원 5건 변환경로 진단',
  enabled: true,
  inputSheet: 'ISSUE43_원천완전성진단',
  outputSheet: 'ISSUE44_매입금액0원추적',
  statusSheet: 'ISSUE44_진단상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var stateSheet = issue44EnsureSheet_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue44Write_(stateSheet, [
    ['항목','값'],
    ['버전','v1.0-ISSUE44-PURCHASE-ZERO-TRACE'],
    ['상태','RUNNING'],
    ['단계','LOAD'],
    ['메시지','매입금액 0원 5건 변환경로 진단 시작'],
    ['운영시트 변경','0'],
    ['갱신시각',new Date().toISOString()]
  ]);

  try {
    var input = ss.getSheetByName(LOTTEON_REMOTE_TASK.inputSheet);
    var source = ss.getSheetByName('매출데이터_붙여넣기');
    var vat = ss.getSheetByName('부가세_신고자료');
    if (!input || input.getLastRow() < 2) throw new Error('ISSUE43_원천완전성진단 시트를 찾지 못했거나 데이터가 없습니다.');
    if (!source || source.getLastRow() < 2) throw new Error('매출데이터_붙여넣기 시트를 찾지 못했거나 데이터가 없습니다.');
    if (!vat || vat.getLastRow() < 2) throw new Error('부가세_신고자료 시트를 찾지 못했거나 데이터가 없습니다.');

    var iv = input.getDataRange().getValues();
    var ih = issue44HeaderMap_(iv[0]);
    var requiredInput = ['주문번호','쿠팡계정ID','원천완전성분류'];
    var missingInput = requiredInput.filter(function(n){ return ih[issue44Norm_(n)] == null; });
    if (missingInput.length) throw new Error('ISSUE43 필수 헤더 누락: ' + missingInput.join(', '));

    var targets = [];
    for (var r = 1; r < iv.length; r++) {
      if (issue44Text_(iv[r][ih[issue44Norm_('원천완전성분류')]]) !== '원천에는 매입금액 있으나 부가세 시트 0원') continue;
      targets.push({
        orderNo: issue44Text_(iv[r][ih[issue44Norm_('주문번호')]]),
        account: issue44Text_(iv[r][ih[issue44Norm_('쿠팡계정ID')]]),
        orderDate: ih[issue44Norm_('주문일')] == null ? '' : issue44DateText_(iv[r][ih[issue44Norm_('주문일')]]),
        issue43Row: r + 1,
        sourceRowsText: ih[issue44Norm_('매출원본행')] == null ? '' : issue44Text_(iv[r][ih[issue44Norm_('매출원본행')]]),
        vatRowsText: ih[issue44Norm_('부가세원본행')] == null ? '' : issue44Text_(iv[r][ih[issue44Norm_('부가세원본행')]])
      });
    }
    if (targets.length !== 5) throw new Error('대상 건수 불일치: 기대 5건, 실제 ' + targets.length + '건');

    var sv = source.getDataRange().getValues();
    var sh = sv[0] || [];
    var sm = issue44HeaderMap_(sh);
    var sourceOrderIx = issue44FindIndex_(sm, ['마켓주문번호','주문번호','주문ID','주문ID(마켓)']);
    var sourceAccountIx = issue44FindIndex_(sm, ['마켓아이디','쿠팡계정ID','계정ID']);
    if (sourceOrderIx < 0) throw new Error('매출데이터_붙여넣기 주문번호 헤더를 찾지 못했습니다.');
    var sourcePurchaseIxs = issue44FindAllIndexes_(sh, ['매입금액','구매가격','매입가격','상품매입금액','결제금액(매입)']);
    var acIx = sh.length > 28 ? 28 : -1;
    var namedPurchaseIxs = sourcePurchaseIxs.filter(function(ix){ return ix !== acIx; });

    var vv = vat.getDataRange().getValues();
    var vh = vv[0] || [];
    var vm = issue44HeaderMap_(vh);
    var vatOrderIx = issue44FindIndex_(vm, ['주문번호','마켓주문번호','주문ID','주문ID(마켓)']);
    var vatAccountIx = issue44FindIndex_(vm, ['쿠팡계정ID','마켓아이디','계정ID']);
    if (vatOrderIx < 0) throw new Error('부가세_신고자료 주문번호 헤더를 찾지 못했습니다.');
    var vatPurchaseIxs = issue44FindAllIndexes_(vh, ['매입금액','구매가격','매입가격','상품매입금액']);
    if (!vatPurchaseIxs.length) throw new Error('부가세_신고자료 매입금액 헤더를 찾지 못했습니다.');

    var outputHeaders = [
      'Issue43행','주문일','쿠팡계정ID','주문번호','원천주문행수','원천계정일치행수','VAT주문행수','VAT계정일치행수',
      'AC열번호','AC헤더','AC합계','AC값상세','매입별칭열','매입별칭합계','매입별칭값상세','VAT매입열','VAT매입합계','VAT값상세',
      '원천행번호','VAT행번호','원천다중행혼재','원인분류','진단메모'
    ];
    var out = [];
    var counts = {};

    targets.forEach(function(t) {
      var sourceOrderRows = issue44FindRows_(sv, sourceOrderIx, t.orderNo);
      var sourceExactRows = sourceAccountIx >= 0 && t.account
        ? sourceOrderRows.filter(function(e){ return issue44Text_(e.row[sourceAccountIx]).toLowerCase() === t.account.toLowerCase(); })
        : sourceOrderRows.slice();
      var sourceRows = sourceExactRows.length ? sourceExactRows : sourceOrderRows;

      var vatOrderRows = issue44FindRows_(vv, vatOrderIx, t.orderNo);
      var vatExactRows = vatAccountIx >= 0 && t.account
        ? vatOrderRows.filter(function(e){ return issue44Text_(e.row[vatAccountIx]).toLowerCase() === t.account.toLowerCase(); })
        : vatOrderRows.slice();
      var vatRows = vatExactRows.length ? vatExactRows : vatOrderRows;

      var acValues = acIx >= 0 ? issue44Values_(sourceRows, [acIx]) : [];
      var namedValues = issue44Values_(sourceRows, namedPurchaseIxs);
      var vatValues = issue44Values_(vatRows, vatPurchaseIxs);
      var acSum = issue44SumValues_(acValues);
      var namedSum = issue44SumValues_(namedValues);
      var vatSum = issue44SumValues_(vatValues);
      var mixed = issue44Mixed_(sourceRows, acIx, namedPurchaseIxs);

      var classification = '';
      var memo = [];
      if (!sourceOrderRows.length) {
        classification = '원천 주문행 누락';
      } else if (sourceAccountIx >= 0 && t.account && !sourceExactRows.length) {
        classification = '주문번호·계정 매칭 불일치';
        memo.push('원천에는 주문번호가 있으나 동일 계정 행이 없음');
      } else if (!vatOrderRows.length) {
        classification = 'VAT 상세행 누락';
      } else if (vatAccountIx >= 0 && t.account && !vatExactRows.length) {
        classification = '주문번호·계정 매칭 불일치';
        memo.push('VAT에는 주문번호가 있으나 동일 계정 행이 없음');
      } else if (mixed) {
        classification = '원천 다중행 혼재';
      } else if (acSum > 0 && vatSum === 0) {
        classification = 'AC 양수 / VAT 0';
      } else if (acSum === 0 && namedSum > 0 && vatSum === 0) {
        classification = '별칭만 양수 / AC 0';
      } else {
        classification = '기타 추가 추적 필요';
      }
      if (acIx >= 0) memo.push('AC=' + issue44ColName_(acIx + 1) + '(' + issue44Text_(sh[acIx]) + ')');
      memo.push('AC합계=' + acSum);
      memo.push('별칭합계=' + namedSum);
      memo.push('VAT합계=' + vatSum);
      if (sourceOrderRows.length !== sourceExactRows.length && sourceAccountIx >= 0 && t.account) memo.push('원천 주문번호 전체행=' + sourceOrderRows.length + ', 계정일치=' + sourceExactRows.length);
      if (vatOrderRows.length !== vatExactRows.length && vatAccountIx >= 0 && t.account) memo.push('VAT 주문번호 전체행=' + vatOrderRows.length + ', 계정일치=' + vatExactRows.length);

      counts[classification] = (counts[classification] || 0) + 1;
      out.push([
        t.issue43Row,t.orderDate,t.account,t.orderNo,sourceOrderRows.length,sourceExactRows.length,vatOrderRows.length,vatExactRows.length,
        acIx >= 0 ? issue44ColName_(acIx + 1) : '',acIx >= 0 ? issue44Text_(sh[acIx]) : '',acSum,issue44ValueSummary_(acValues),
        issue44IndexSummary_(sh,namedPurchaseIxs),namedSum,issue44ValueSummary_(namedValues),issue44IndexSummary_(vh,vatPurchaseIxs),vatSum,issue44ValueSummary_(vatValues),
        issue44RowNos_(sourceRows),issue44RowNos_(vatRows),mixed ? 'Y' : 'N',classification,memo.join(' / ')
      ]);
    });

    if (out.length !== 5) throw new Error('출력 건수 불일치: ' + out.length);
    var output = issue44EnsureSheet_(ss, LOTTEON_REMOTE_TASK.outputSheet);
    output.clearContents();
    output.getRange(1,1,1,outputHeaders.length).setValues([outputHeaders]);
    output.getRange(2,1,out.length,outputHeaders.length).setValues(out);
    output.setFrozenRows(1);
    output.getRange(1,1,1,outputHeaders.length).setBackground('#d9eaf7').setFontWeight('bold');
    output.getRange(2,11,out.length,1).setNumberFormat('#,##0');
    output.getRange(2,14,out.length,1).setNumberFormat('#,##0');
    output.getRange(2,17,out.length,1).setNumberFormat('#,##0');

    var statusRows = [
      ['항목','값'],
      ['버전','v1.0-ISSUE44-PURCHASE-ZERO-TRACE'],
      ['상태','PASS'],
      ['단계','DONE'],
      ['메시지','매입금액 0원 5건 변환경로 진단 완료'],
      ['대상건수',targets.length],
      ['출력건수',out.length],
      ['운영시트 변경','0']
    ];
    Object.keys(counts).sort().forEach(function(k){ statusRows.push(['분류_' + k, counts[k]]); });
    statusRows.push(['완료시각',new Date().toISOString()]);
    issue44Write_(stateSheet,statusRows);
    try {
      MailApp.sendEmail('beliun1001@gmail.com','[LOTTEON 자동작업 결과][PASS] ISSUE44-v1.0',statusRows.map(function(x){return x[0]+': '+x[1];}).join('\n'));
    } catch (mailError) {
      statusRows.push(['완료알림오류',String(mailError && mailError.message ? mailError.message : mailError)]);
      issue44Write_(stateSheet,statusRows);
    }
    return {ok:true,target:targets.length,counts:counts};
  } catch (e) {
    issue44Write_(stateSheet,[
      ['항목','값'],
      ['버전','v1.0-ISSUE44-PURCHASE-ZERO-TRACE'],
      ['상태','ERROR'],
      ['단계','FAILED'],
      ['메시지','매입금액 0원 5건 변환경로 진단 실패'],
      ['오류',String(e && e.message ? e.message : e)],
      ['운영시트 변경','0'],
      ['갱신시각',new Date().toISOString()]
    ]);
    throw e;
  }
}

function issue44FindRows_(values, orderIx, orderNo) {
  var out=[];
  for (var r=1;r<values.length;r++) if (issue44Text_(values[r][orderIx])===orderNo) out.push({rowNo:r+1,row:values[r]});
  return out;
}
function issue44Values_(entries,indexes) {
  var out=[];
  (entries||[]).forEach(function(e){(indexes||[]).forEach(function(ix){out.push({rowNo:e.rowNo,index:ix,value:e.row[ix],number:issue44Number_(e.row[ix])});});});
  return out;
}
function issue44SumValues_(values){var n=0;(values||[]).forEach(function(v){n+=v.number;});return n;}
function issue44Mixed_(entries,acIx,namedIxs){
  if (!entries || entries.length < 2) return false;
  var vals=[];
  entries.forEach(function(e){
    var candidates=[];
    if(acIx>=0)candidates.push(issue44Number_(e.row[acIx]));
    (namedIxs||[]).forEach(function(ix){candidates.push(issue44Number_(e.row[ix]));});
    var max=0;candidates.forEach(function(n){if(n>max)max=n;});vals.push(max);
  });
  var hasPos=vals.some(function(n){return n>0;});
  var hasZero=vals.some(function(n){return n===0;});
  var uniq={};vals.forEach(function(n){uniq[String(n)]=true;});
  return hasPos && (hasZero || Object.keys(uniq).length>1);
}
function issue44ValueSummary_(values){return (values||[]).map(function(v){return 'R'+v.rowNo+':'+issue44ColName_(v.index+1)+'='+issue44Text_(v.value);}).join(' | ');}
function issue44IndexSummary_(headers,indexes){return (indexes||[]).map(function(ix){return issue44ColName_(ix+1)+'('+issue44Text_(headers[ix])+')';}).join(' | ');}
function issue44RowNos_(entries){return (entries||[]).map(function(e){return e.rowNo;}).join(',');}
function issue44FindAllIndexes_(headers,aliases){var set={};aliases.forEach(function(a){set[issue44Norm_(a)]=true;});var out=[];(headers||[]).forEach(function(h,i){if(set[issue44Norm_(h)])out.push(i);});return out;}
function issue44FindIndex_(map,aliases){for(var i=0;i<aliases.length;i++){var k=issue44Norm_(aliases[i]);if(map[k]!=null)return map[k];}return -1;}
function issue44HeaderMap_(headers){var map={};(headers||[]).forEach(function(h,i){map[issue44Norm_(h)]=i;});return map;}
function issue44Norm_(v){return issue44Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function issue44Text_(v){return v==null?'':String(v).trim();}
function issue44Number_(v){if(typeof v==='number'&&isFinite(v))return v;var n=Number(String(v==null?'':v).replace(/,/g,'').replace(/[^0-9.\-]/g,''));return isFinite(n)?n:0;}
function issue44DateText_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');return issue44Text_(v);}
function issue44ColName_(n){var s='';while(n>0){var m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;}
function issue44EnsureSheet_(ss,name){return ss.getSheetByName(name)||ss.insertSheet(name);}
function issue44Write_(sheet,rows){sheet.clearContents();sheet.getRange(1,1,rows.length,2).setValues(rows);sheet.setFrozenRows(1);sheet.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');SpreadsheetApp.flush();}
function runLotteonRemoteTaskContinueRemote_(){return runLotteonRemoteTaskStartRemote_();}
