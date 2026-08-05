/**
 * Issue #42 v1.1 standalone operating diagnostic.
 * Reads only production VAT/card evidence sheets and writes only ISSUE42_* sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE42-v1.1-20260805',
  title: '상반기 VAT 잔여 51건 원인 분류 진단',
  enabled: true,
  statusSheet: 'ISSUE42_진단상태',
  outputSheet: 'ISSUE42_잔여매칭진단'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var statusSheet = issue42EnsureSheet_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue42WriteStatus_(statusSheet, [
    ['항목','값'],
    ['버전','v1.1-ISSUE42-STANDALONE-DIAGNOSTIC'],
    ['상태','RUNNING'],
    ['단계','LOAD'],
    ['메시지','잔여 카드매칭 진단 시작'],
    ['운영시트 변경','0'],
    ['갱신시각',new Date().toISOString()]
  ]);

  try {
    var diagSheet = ss.getSheetByName('부가세_카드매칭검증');
    var historySheet = ss.getSheetByName('카드사용내역_붙여넣기');
    if (!diagSheet || diagSheet.getLastRow() < 2) throw new Error('부가세_카드매칭검증 시트를 찾지 못했거나 데이터가 없습니다.');
    if (!historySheet || historySheet.getLastRow() < 2) throw new Error('카드사용내역_붙여넣기 시트를 찾지 못했거나 데이터가 없습니다.');

    var diagValues = diagSheet.getDataRange().getValues();
    var dh = issue42HeaderMap_(diagValues[0]);
    var required = ['신고연도','반기','주문일','사업자등록번호','쿠팡계정ID','주문번호','롯데결제수단','주문매입금액','카드매칭상태','카드매칭근거'];
    var missing = required.filter(function(name){ return dh[name] == null; });
    if (missing.length) throw new Error('부가세_카드매칭검증 필수 헤더 누락: ' + missing.join(', '));

    var usedApprovalKeys = {};
    var unresolved = [];
    for (var r = 1; r < diagValues.length; r++) {
      var row = diagValues[r];
      var status = issue42Text_(row[dh['카드매칭상태']]).toUpperCase();
      var company = dh['구매카드사'] == null ? '' : issue42Text_(row[dh['구매카드사']]);
      var approvalNo = dh['승인번호'] == null ? '' : issue42Text_(row[dh['승인번호']]);
      if ((status === 'MATCHED' || status === 'MASTER_MATCHED' || status === 'NON_CARD') && approvalNo) {
        usedApprovalKeys[issue42ApprovalKey_(company, approvalNo)] = true;
      }
      if (issue42Text_(row[dh['신고연도']]) === '2026' && issue42Text_(row[dh['반기']]) === '상반기' &&
          (status === 'NO_MATCH' || status === 'AMBIGUOUS')) {
        unresolved.push({
          sourceRow:r + 1,
          orderDate:issue42DateText_(row[dh['주문일']]),
          business:issue42Text_(row[dh['사업자등록번호']]),
          account:issue42Text_(row[dh['쿠팡계정ID']]),
          orderNo:issue42Text_(row[dh['주문번호']]),
          payment:issue42Text_(row[dh['롯데결제수단']]),
          purchase:issue42Number_(row[dh['주문매입금액']]),
          currentStatus:status,
          currentReason:issue42Text_(row[dh['카드매칭근거']]),
          currentCandidateCount:dh['후보수'] == null ? 0 : issue42Number_(row[dh['후보수']])
        });
      }
    }
    if (unresolved.length !== 51) throw new Error('대상 건수 불일치: 기대 51건, 실제 ' + unresolved.length + '건');

    var historyValues = historySheet.getDataRange().getValues();
    var hh = issue42HeaderMap_(historyValues[0]);
    var historyRequired = ['카드사','승인일','승인금액','승인번호'];
    var missingHistory = historyRequired.filter(function(name){ return hh[name] == null; });
    if (missingHistory.length) throw new Error('카드사용내역_붙여넣기 필수 헤더 누락: ' + missingHistory.join(', '));

    var evidence = [];
    for (var i = 1; i < historyValues.length; i++) {
      var hrow = historyValues[i];
      var company2 = issue42Text_(hrow[hh['카드사']]);
      var date = issue42DateText_(hrow[hh['승인일']]);
      var amount = issue42Number_(hrow[hh['승인금액']]);
      var approval = issue42Text_(hrow[hh['승인번호']]);
      var status2 = hh['승인상태'] == null ? '' : issue42Text_(hrow[hh['승인상태']]);
      var merchant = hh['가맹점명'] == null ? '' : issue42Text_(hrow[hh['가맹점명']]);
      var lotteFlag = hh['롯데계열여부'] == null ? '' : issue42Text_(hrow[hh['롯데계열여부']]);
      var cancelAmount = hh['취소금액'] == null ? 0 : Math.abs(issue42Number_(hrow[hh['취소금액']]));
      var cardName = hh['카드명'] == null ? '' : issue42Text_(hrow[hh['카드명']]);
      var cardNo = hh['카드번호'] == null ? '' : issue42Text_(hrow[hh['카드번호']]);
      var end4 = hh['카드번호끝4'] == null ? '' : issue42PadEnd4_(hrow[hh['카드번호끝4']]);
      var merchantOrderNo = hh['가맹점주문번호'] == null ? '' : issue42Text_(hrow[hh['가맹점주문번호']]);
      var sourceFile = hh['원본파일'] == null ? '' : issue42Text_(hrow[hh['원본파일']]);
      var evidenceType = hh['증빙유형'] == null ? '' : issue42Text_(hrow[hh['증빙유형']]);
      var isCancel = /취소|환불|반품/.test(status2) || amount < 0;
      var effective = Math.max(Math.abs(amount) - cancelAmount, 0);
      var lotte = /^(y|yes|true|1|롯데)$/i.test(lotteFlag) || /롯데|lotte/i.test(merchant);
      if (!date || !effective || isCancel || !lotte) continue;
      evidence.push({
        rowNo:i + 1,
        company:company2,
        cardName:cardName,
        cardNumber:cardNo,
        cardEnd4:end4,
        date:date,
        amount:effective,
        approvalNo:approval,
        merchant:merchant,
        merchantOrderNo:merchantOrderNo,
        sourceFile:sourceFile,
        evidenceType:evidenceType,
        used:approval ? !!usedApprovalKeys[issue42ApprovalKey_(company2, approval)] : false
      });
    }

    var outHeaders = [
      '원본행','주문일','사업자등록번호','쿠팡계정ID','주문번호','롯데결제수단','주문매입금액',
      '현재상태','현재근거','현재후보수','주원인분류','추가검토가능','0~+7일후보','+8~+14일후보',
      '이전1~7일후보','기타±30일후보','미사용후보','사용된후보','후보카드수','후보카드요약','후보증빙요약'
    ];
    var outRows = [];
    var counts = {};
    unresolved.forEach(function(order) {
      var exact = evidence.filter(function(ev){ return ev.amount === order.purchase; });
      var w0 = [], w8 = [], prev = [], other = [];
      exact.forEach(function(ev) {
        var d = issue42DaysBetween_(order.orderDate, ev.date);
        if (d >= 0 && d <= 7) w0.push(ev);
        else if (d >= 8 && d <= 14) w8.push(ev);
        else if (d >= -7 && d <= -1) prev.push(ev);
        else if (Math.abs(d) <= 30) other.push(ev);
      });
      var allWindow = w0.concat(w8, prev, other);
      var unused = allWindow.filter(function(ev){ return !ev.used; });
      var used = allWindow.filter(function(ev){ return ev.used; });
      var identities = {};
      unused.forEach(function(ev){ identities[issue42IdentityKey_(ev)] = true; });
      var identityCount = Object.keys(identities).length;
      var classification = issue42Classify_(order, w0, w8, prev, other, unused, used);
      counts[classification] = (counts[classification] || 0) + 1;
      var reviewable = /단일후보|동일카드 다중건/.test(classification) ? 'Y' : 'N';
      outRows.push([
        order.sourceRow,order.orderDate,order.business,order.account,order.orderNo,order.payment,order.purchase,
        order.currentStatus,order.currentReason,order.currentCandidateCount,classification,reviewable,
        w0.length,w8.length,prev.length,other.length,unused.length,used.length,identityCount,
        issue42CandidateCardSummary_(unused),issue42EvidenceSummary_(allWindow)
      ]);
    });

    if (outRows.length !== 51) throw new Error('출력 건수 불일치: ' + outRows.length);
    var outputSheet = issue42EnsureSheet_(ss, LOTTEON_REMOTE_TASK.outputSheet);
    outputSheet.clearContents();
    outputSheet.getRange(1,1,1,outHeaders.length).setValues([outHeaders]);
    outputSheet.getRange(2,1,outRows.length,outHeaders.length).setValues(outRows);
    outputSheet.setFrozenRows(1);
    outputSheet.getRange(1,1,1,outHeaders.length).setBackground('#d9eaf7').setFontWeight('bold');
    outputSheet.getRange(2,7,outRows.length,1).setNumberFormat('#,##0');

    var statusRows = [
      ['항목','값'],
      ['버전','v1.1-ISSUE42-STANDALONE-DIAGNOSTIC'],
      ['상태','PASS'],
      ['단계','DONE'],
      ['메시지','잔여 카드매칭 원인 분류 완료'],
      ['대상건수',unresolved.length],
      ['출력건수',outRows.length],
      ['NO_MATCH',unresolved.filter(function(x){ return x.currentStatus === 'NO_MATCH'; }).length],
      ['AMBIGUOUS',unresolved.filter(function(x){ return x.currentStatus === 'AMBIGUOUS'; }).length],
      ['추가검토가능',outRows.filter(function(x){ return x[11] === 'Y'; }).length],
      ['운영시트 변경','0']
    ];
    Object.keys(counts).sort().forEach(function(key){ statusRows.push(['분류_' + key, counts[key]]); });
    statusRows.push(['완료시각',new Date().toISOString()]);
    issue42WriteStatus_(statusSheet,statusRows);

    try {
      MailApp.sendEmail('beliun1001@gmail.com','[LOTTEON 자동작업 결과][PASS] ISSUE42-v1.1',
        statusRows.map(function(x){ return x[0] + ': ' + x[1]; }).join('\n'));
    } catch (mailError) {
      statusRows.push(['완료알림오류',String(mailError && mailError.message ? mailError.message : mailError)]);
      issue42WriteStatus_(statusSheet,statusRows);
    }
    return {ok:true,target:unresolved.length,output:outRows.length,counts:counts};
  } catch (e) {
    issue42WriteStatus_(statusSheet,[
      ['항목','값'],
      ['버전','v1.1-ISSUE42-STANDALONE-DIAGNOSTIC'],
      ['상태','ERROR'],
      ['단계','FAILED'],
      ['메시지','잔여 카드매칭 진단 실패'],
      ['오류',String(e && e.message ? e.message : e)],
      ['운영시트 변경','0'],
      ['갱신시각',new Date().toISOString()]
    ]);
    throw e;
  }
}

function issue42Classify_(order,w0,w8,prev,other,unused,used) {
  if (!order.purchase) return '금액 0원';
  if (!order.payment) return '결제수단 공란';
  var w0u = w0.filter(function(x){return !x.used;});
  var w8u = w8.filter(function(x){return !x.used;});
  var prevu = prev.filter(function(x){return !x.used;});
  if (w0u.length) {
    var n0 = issue42IdentityCount_(w0u);
    return n0 > 1 ? 'exact 후보 다중카드' : (w0u.length > 1 ? 'exact 후보 동일카드 다중건' : '0~+7일 미사용 단일후보');
  }
  if (w8u.length && issue42IdentityCount_(w8u) === 1) return w8u.length === 1 ? '+8~+14일 단일후보' : '+8~+14일 동일카드 다중건';
  if (prevu.length && issue42IdentityCount_(prevu) === 1) return prevu.length === 1 ? '주문일 이전 단일후보' : '주문일 이전 동일카드 다중건';
  if (unused.length && issue42IdentityCount_(unused) > 1) return '±30일 exact 후보 다중카드';
  if (unused.length) return '±30일 기타 단일카드 후보';
  if (used.length) return '증빙은 있으나 이미 다른 주문에 사용됨';
  if (order.currentStatus === 'AMBIGUOUS') return '현재 AMBIGUOUS 유지 필요';
  return 'exact 후보 없음';
}
function issue42IdentityCount_(rows){var m={};rows.forEach(function(x){m[issue42IdentityKey_(x)]=true;});return Object.keys(m).length;}
function issue42IdentityKey_(ev){return [issue42NormCompany_(ev.company),ev.cardEnd4||'',ev.cardNumber||'',ev.cardName||''].join('|');}
function issue42NormCompany_(v){return issue42Text_(v).toLowerCase().replace(/\s|카드/g,'');}
function issue42ApprovalKey_(company,no){return issue42NormCompany_(company)+'|'+issue42Text_(no);}
function issue42CandidateCardSummary_(rows){var m={};rows.forEach(function(x){var label=[x.company,x.cardName,x.cardEnd4].filter(Boolean).join('/');m[label||'식별불가']=true;});return Object.keys(m).sort().join(' ; ');}
function issue42EvidenceSummary_(rows){return rows.slice(0,8).map(function(x){return [x.date,x.company,x.cardEnd4,x.approvalNo,x.used?'USED':'UNUSED'].filter(Boolean).join('/')}).join(' ; ');}
function issue42PadEnd4_(v){var s=issue42Text_(v).replace(/\D/g,'');return s?s.slice(-4).padStart(4,'0'):'';}
function issue42HeaderMap_(row){var m={};(row||[]).forEach(function(v,i){m[issue42Text_(v)]=i;});return m;}
function issue42Text_(v){return v == null ? '' : String(v).trim();}
function issue42Number_(v){if(typeof v==='number')return isFinite(v)?v:0;var n=Number(String(v==null?'':v).replace(/,/g,'').trim());return isFinite(n)?n:0;}
function issue42DateText_(v){
  if (Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime())) return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');
  var s=issue42Text_(v);var m=s.match(/(20\d{2})[-\/.년\s]*(\d{1,2})[-\/.월\s]*(\d{1,2})/);if(!m)return s;return m[1]+'-'+String(Number(m[2])).padStart(2,'0')+'-'+String(Number(m[3])).padStart(2,'0');
}
function issue42DaysBetween_(a,b){var x=issue42DateText_(a).match(/^(\d{4})-(\d{2})-(\d{2})$/),y=issue42DateText_(b).match(/^(\d{4})-(\d{2})-(\d{2})$/);if(!x||!y)return 99999;return Math.round((Date.UTC(+y[1],+y[2]-1,+y[3])-Date.UTC(+x[1],+x[2]-1,+x[3]))/86400000);}
function issue42EnsureSheet_(ss,name){return ss.getSheetByName(name)||ss.insertSheet(name);}
function issue42WriteStatus_(sheet,rows){sheet.clearContents();sheet.getRange(1,1,rows.length,2).setValues(rows);sheet.setFrozenRows(1);sheet.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');}
function runLotteonRemoteTaskContinueRemote_(){return runLotteonRemoteTaskStartRemote_();}
