/**
 * Issue #43 standalone source-completeness diagnostic.
 * Reads ISSUE42 output + production VAT/source sheets and writes only ISSUE43_* sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE43-v1.0-20260805',
  title: '상반기 VAT 잔여 51건 원천 완전성 진단',
  enabled: true,
  inputSheet: 'ISSUE42_잔여매칭진단',
  outputSheet: 'ISSUE43_원천완전성진단',
  statusSheet: 'ISSUE43_진단상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var stateSheet = issue43EnsureSheet_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue43Write_(stateSheet, [
    ['항목','값'],
    ['버전','v1.0-ISSUE43-SOURCE-COMPLETENESS-DIAGNOSTIC'],
    ['상태','RUNNING'],
    ['단계','LOAD'],
    ['메시지','원천 완전성 진단 시작'],
    ['운영시트 변경','0'],
    ['갱신시각',new Date().toISOString()]
  ]);

  try {
    var issue42Sheet = ss.getSheetByName(LOTTEON_REMOTE_TASK.inputSheet);
    var vatSheet = ss.getSheetByName('부가세_신고자료');
    var sourceSheet = ss.getSheetByName('매출데이터_붙여넣기');
    if (!issue42Sheet || issue42Sheet.getLastRow() < 2) throw new Error('ISSUE42_잔여매칭진단 시트를 찾지 못했거나 데이터가 없습니다.');
    if (!vatSheet || vatSheet.getLastRow() < 2) throw new Error('부가세_신고자료 시트를 찾지 못했거나 데이터가 없습니다.');
    if (!sourceSheet || sourceSheet.getLastRow() < 2) throw new Error('매출데이터_붙여넣기 시트를 찾지 못했거나 데이터가 없습니다.');

    var targetValues = issue42Sheet.getDataRange().getValues();
    var th = issue43HeaderMap_(targetValues[0]);
    var targetRequired = ['주문번호','주문일','쿠팡계정ID','롯데결제수단','주문매입금액','주원인분류'];
    var targetMissing = targetRequired.filter(function(name){ return th[name] == null; });
    if (targetMissing.length) throw new Error('ISSUE42 필수 헤더 누락: ' + targetMissing.join(', '));

    var vatValues = vatSheet.getDataRange().getValues();
    var vh = issue43HeaderMap_(vatValues[0]);
    var vatOrderIndex = issue43FindIndex_(vh, ['주문번호','마켓주문번호','주문ID','주문ID(마켓)']);
    if (vatOrderIndex < 0) throw new Error('부가세_신고자료 주문번호 헤더를 찾지 못했습니다.');
    var vatAccountIndex = issue43FindIndex_(vh, ['쿠팡계정ID','마켓아이디','계정ID']);
    var vatPaymentIndexes = issue43FindAllIndexes_(vatValues[0], ['롯데결제수단','구매결제수단','결제수단','결제정보','결제방법','카드사','결제수단/카드사','결제수단(카드사)']);
    var vatPurchaseIndexes = issue43FindAllIndexes_(vatValues[0], ['매입금액','구매가격','매입가격','상품매입금액']);

    var sourceValues = sourceSheet.getDataRange().getValues();
    var sh = issue43HeaderMap_(sourceValues[0]);
    var sourceOrderIndex = issue43FindIndex_(sh, ['마켓주문번호','주문번호','주문ID','주문ID(마켓)']);
    if (sourceOrderIndex < 0) throw new Error('매출데이터_붙여넣기 주문번호 헤더를 찾지 못했습니다.');
    var sourceAccountIndex = issue43FindIndex_(sh, ['마켓아이디','쿠팡계정ID','계정ID']);
    var sourcePaymentIndexes = issue43FindAllIndexes_(sourceValues[0], ['롯데결제수단','구매결제수단','결제수단','결제정보','결제방법','카드사','결제수단/카드사','결제수단(카드사)']);
    var sourcePurchaseIndexes = issue43FindAllIndexes_(sourceValues[0], ['매입금액','구매가격','매입가격','상품매입금액','결제금액(매입)']);
    if (sourcePurchaseIndexes.indexOf(28) < 0 && sourceValues[0].length > 28) sourcePurchaseIndexes.push(28); // AC source-of-truth
    var sourceStatusIndexes = issue43FindAllIndexes_(sourceValues[0], ['마켓주문상태','주문상태','상태','클레임상태','처리상태']);

    var vatMap = issue43BuildOrderMap_(vatValues, vatOrderIndex, vatAccountIndex);
    var sourceMap = issue43BuildOrderMap_(sourceValues, sourceOrderIndex, sourceAccountIndex);

    var outputHeaders = [
      '주문일','쿠팡계정ID','주문번호','ISSUE42분류','현재롯데결제수단','현재주문매입금액',
      '부가세상세행수','원천행수','부가세결제수단','원천결제수단','부가세매입금액합계','원천AC매입금액합계',
      '원천매입금액값','원천주문상태','부가세원본행','매출원본행','원천완전성분류','진단메모'
    ];
    var outputRows = [];
    var counts = {};

    for (var r = 1; r < targetValues.length; r++) {
      var target = targetValues[r];
      var orderNo = issue43Text_(target[th['주문번호']]);
      if (!orderNo) continue;
      var account = issue43Text_(target[th['쿠팡계정ID']]);
      var issue42Class = issue43Text_(target[th['주원인분류']]);
      var currentPayment = issue43Text_(target[th['롯데결제수단']]);
      var currentPurchase = issue43Number_(target[th['주문매입금액']]);
      var key = issue43OrderKey_(orderNo, account);
      var vatEntries = vatMap[key] || vatMap[issue43OrderKey_(orderNo, '')] || [];
      var sourceEntries = sourceMap[key] || sourceMap[issue43OrderKey_(orderNo, '')] || [];

      var vatPayments = issue43CollectTexts_(vatEntries, vatPaymentIndexes);
      var sourcePayments = issue43CollectTexts_(sourceEntries, sourcePaymentIndexes);
      var vatPurchase = issue43SumIndexes_(vatEntries, vatPurchaseIndexes);
      var sourcePurchase = issue43SumIndexes_(sourceEntries, sourcePurchaseIndexes);
      var sourcePurchaseValues = issue43CollectNumbers_(sourceEntries, sourcePurchaseIndexes);
      var statuses = issue43CollectTexts_(sourceEntries, sourceStatusIndexes);
      var result = issue43Classify_(issue42Class, sourceEntries.length, sourcePayments, sourcePurchase, statuses);
      counts[result.classification] = (counts[result.classification] || 0) + 1;

      outputRows.push([
        issue43DateText_(target[th['주문일']]),account,orderNo,issue42Class,currentPayment,currentPurchase,
        vatEntries.length,sourceEntries.length,vatPayments.join(' | '),sourcePayments.join(' | '),vatPurchase,sourcePurchase,
        sourcePurchaseValues.join(' | '),statuses.join(' | '),issue43RowNos_(vatEntries),issue43RowNos_(sourceEntries),
        result.classification,result.memo
      ]);
    }

    if (outputRows.length !== 51) throw new Error('대상 건수 불일치: 기대 51건, 실제 ' + outputRows.length + '건');
    var outputSheet = issue43EnsureSheet_(ss, LOTTEON_REMOTE_TASK.outputSheet);
    outputSheet.clearContents();
    outputSheet.getRange(1,1,1,outputHeaders.length).setValues([outputHeaders]);
    outputSheet.getRange(2,1,outputRows.length,outputHeaders.length).setValues(outputRows);
    outputSheet.setFrozenRows(1);
    outputSheet.getRange(1,1,1,outputHeaders.length).setBackground('#d9eaf7').setFontWeight('bold');
    outputSheet.getRange(2,6,outputRows.length,1).setNumberFormat('#,##0');
    outputSheet.getRange(2,11,outputRows.length,2).setNumberFormat('#,##0');

    var statusRows = [
      ['항목','값'],
      ['버전','v1.0-ISSUE43-SOURCE-COMPLETENESS-DIAGNOSTIC'],
      ['상태','PASS'],
      ['단계','DONE'],
      ['메시지','원천 결제수단·매입금액 완전성 진단 완료'],
      ['대상건수',outputRows.length],
      ['운영시트 변경','0']
    ];
    Object.keys(counts).sort().forEach(function(name){ statusRows.push(['분류_' + name, counts[name]]); });
    statusRows.push(['완료시각',new Date().toISOString()]);
    issue43Write_(stateSheet,statusRows);
    try {
      MailApp.sendEmail('beliun1001@gmail.com','[LOTTEON 자동작업 결과][PASS] ISSUE43-v1.0',
        statusRows.map(function(row){ return row[0] + ': ' + row[1]; }).join('\n'));
    } catch (mailError) {
      statusRows.push(['완료알림오류',String(mailError && mailError.message ? mailError.message : mailError)]);
      issue43Write_(stateSheet,statusRows);
    }
    return {ok:true,target:outputRows.length,counts:counts};
  } catch (e) {
    issue43Write_(stateSheet,[
      ['항목','값'],
      ['버전','v1.0-ISSUE43-SOURCE-COMPLETENESS-DIAGNOSTIC'],
      ['상태','ERROR'],
      ['단계','FAILED'],
      ['메시지','원천 완전성 진단 실패'],
      ['오류',String(e && e.message ? e.message : e)],
      ['운영시트 변경','0'],
      ['갱신시각',new Date().toISOString()]
    ]);
    throw e;
  }
}

function issue43Classify_(issue42Class, sourceCount, sourcePayments, sourcePurchase, statuses) {
  if (!sourceCount) return {classification:'원천 주문행 없음',memo:'매출데이터_붙여넣기에서 동일 주문번호를 찾지 못함'};
  var statusText = statuses.join(' | ');
  if (issue42Class === '결제수단 공란') {
    if (sourcePayments.length) return {classification:'원천에는 결제수단 있으나 부가세 시트 누락',memo:'원천 결제수단=' + sourcePayments.join(' | ')};
    return {classification:'원천에도 결제수단 공란',memo:'확인한 결제수단 관련 컬럼이 모두 공란'};
  }
  if (issue42Class === '금액 0원') {
    if (/취소|반품|환불|교환/.test(statusText)) return {classification:'취소/반품/환불 관련 0원',memo:'원천상태=' + statusText};
    if (sourcePurchase > 0) return {classification:'원천에는 매입금액 있으나 부가세 시트 0원',memo:'원천 매입금액 합계=' + sourcePurchase};
    return {classification:'AC 매입금액 실제 0원',memo:'원천 AC/매입금액 관련 값 합계가 0'};
  }
  if (issue42Class === 'exact 후보 없음') {
    return {classification:'카드 원본 증빙 자체 없음',memo:'Issue42 ±30일 exact amount 후보 0건'};
  }
  return {classification:'기타 원천 확인 필요',memo:'Issue42 분류=' + issue42Class};
}

function issue43BuildOrderMap_(values, orderIndex, accountIndex) {
  var map = {};
  for (var r = 1; r < values.length; r++) {
    var orderNo = issue43Text_(values[r][orderIndex]);
    if (!orderNo) continue;
    var account = accountIndex >= 0 ? issue43Text_(values[r][accountIndex]) : '';
    var entry = {rowNo:r + 1,row:values[r]};
    var exactKey = issue43OrderKey_(orderNo,account);
    var blankKey = issue43OrderKey_(orderNo,'');
    if (!map[exactKey]) map[exactKey] = [];
    map[exactKey].push(entry);
    if (exactKey !== blankKey) {
      if (!map[blankKey]) map[blankKey] = [];
      map[blankKey].push(entry);
    }
  }
  return map;
}
function issue43OrderKey_(orderNo,account){return issue43Text_(orderNo) + '|' + issue43Text_(account).toLowerCase();}
function issue43RowNos_(entries){return entries.map(function(x){return x.rowNo;}).join(',');}
function issue43CollectTexts_(entries,indexes){
  var out = {}, list = [];
  entries.forEach(function(entry){indexes.forEach(function(index){var value=issue43Text_(entry.row[index]);if(value&&!out[value]){out[value]=true;list.push(value);}});});
  return list;
}
function issue43CollectNumbers_(entries,indexes){
  var list=[];
  entries.forEach(function(entry){indexes.forEach(function(index){var n=issue43Number_(entry.row[index]);if(n||String(entry.row[index]||'').trim()==='0')list.push(n);});});
  return list;
}
function issue43SumIndexes_(entries,indexes){
  var sum=0;
  entries.forEach(function(entry){indexes.forEach(function(index){sum+=issue43Number_(entry.row[index]);});});
  return sum;
}
function issue43FindAllIndexes_(headers,aliases){
  var normalizedAliases={};aliases.forEach(function(x){normalizedAliases[issue43Norm_(x)]=true;});
  var out=[];(headers||[]).forEach(function(header,index){if(normalizedAliases[issue43Norm_(header)])out.push(index);});
  return out;
}
function issue43FindIndex_(map,aliases){for(var i=0;i<aliases.length;i++){var key=issue43Norm_(aliases[i]);if(map[key]!=null)return map[key];}return -1;}
function issue43HeaderMap_(headers){var map={};(headers||[]).forEach(function(x,i){map[issue43Norm_(x)]=i;map[issue43Text_(x)]=i;});return map;}
function issue43Norm_(value){return issue43Text_(value).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function issue43Text_(value){return value==null?'':String(value).trim();}
function issue43Number_(value){if(typeof value==='number'&&isFinite(value))return value;var n=Number(String(value==null?'':value).replace(/,/g,'').replace(/[^0-9.\-]/g,''));return isFinite(n)?n:0;}
function issue43DateText_(value){
  if(Object.prototype.toString.call(value)==='[object Date]'&&!isNaN(value.getTime()))return Utilities.formatDate(value,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');
  var text=issue43Text_(value);var m=text.match(/(20\d{2})[^0-9]?(\d{1,2})[^0-9]?(\d{1,2})/);return m?m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2):text;
}
function issue43EnsureSheet_(ss,name){return ss.getSheetByName(name)||ss.insertSheet(name);}
function issue43Write_(sheet,rows){sheet.clearContents();sheet.getRange(1,1,rows.length,2).setValues(rows);sheet.setFrozenRows(1);sheet.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');SpreadsheetApp.flush();}
function runLotteonRemoteTaskContinueRemote_(){return runLotteonRemoteTaskStartRemote_();}
